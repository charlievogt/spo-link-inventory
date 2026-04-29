#!/usr/bin/env node
// Demo seeder: stale-copy duplicates + sharing-link references.
//
// Stale-copy demo:
//   1. Generate vendor-onboarding.pdf v1 and v2 (different content, same name).
//   2. Upload v1 to /sites/demo-hr (will be the stale copy).
//   3. Upload v1 to /sites/demo-finance.
//   4. Overwrite finance's copy with v2 (creates version history; v1 stays
//      as a prior version).
//   5. Result: HR's current hash == Finance's prior-version hash → HR's copy
//      is detected as stale once the duplicates bootstrap walks history.
//
// Sharing-link demo:
//   1. Generate org-view sharing links for a handful of existing demo files
//      (Steve.pdf, the runbook, a brand asset) via SP REST.
//   2. Create a new modern page on /sites/demo-engineering that embeds those
//      :b:/s/<site>/<token> URLs alongside direct AllItems.aspx links — so
//      the inventory can show both shapes side-by-side.
//
// Sharing-link URLs are the biggest pain point for SP admins because the
// token is opaque (no path), they can outlive the file or be revoked
// silently, and they don't appear in the file's permissions UI.
//
// Usage:
//   node scripts/demo-fixtures/seed-stale-and-sharing.mjs
//
// Reads cert auth from scripts/.entra-output.env.

import { readFileSync } from 'node:fs';
import { createSign, createHash, randomBytes } from 'node:crypto';
import { PDFDocument, StandardFonts, rgb } from 'pdf-lib';

// ---------- 1. Load env ---------------------------------------------------

const env = Object.fromEntries(
  readFileSync(new URL('../.entra-output.env', import.meta.url), 'utf8')
    .split('\n')
    .filter((l) => l && !l.startsWith('#'))
    .map((l) => { const i = l.indexOf('='); return [l.slice(0, i), l.slice(i + 1)]; }),
);
const TENANT_ID = env.ENTRA_TENANT_ID;
const CLIENT_ID = env.ENTRA_CLIENT_ID;
const PEM = Buffer.from(env.ENTRA_CLIENT_CERT_PEM_BASE64, 'base64').toString('utf8');
const CERT_DER = Buffer.from(
  PEM.match(/-----BEGIN CERTIFICATE-----[\s\S]+?-----END CERTIFICATE-----/)[0]
    .replace(/-----BEGIN CERTIFICATE-----|-----END CERTIFICATE-----|\s/g, ''),
  'base64',
);
// deepcode ignore InsecureHash: x5t is a base64url-encoded SHA-1 cert thumbprint per RFC 7515 §4.1.7; Entra rejects client_assertion JWTs that use any other algorithm here.
const X5T = createHash('sha1').update(CERT_DER).digest('base64url');

// HACK: tenant origin isn't in the env file; derive from a known site URL or
// expect SPO_ORIGIN to be set. For now, hardcode contoso — the seeder is for
// the contoso demo tenant.
const SPO_ORIGIN = process.env.SPO_ORIGIN ?? 'https://contoso.sharepoint.com';

// ---------- 2. Cert-based app-only token --------------------------------

async function appOnlyToken() {
  const now = Math.floor(Date.now() / 1000);
  const enc = (o) => Buffer.from(JSON.stringify(o)).toString('base64url');
  const header = enc({ alg: 'RS256', typ: 'JWT', x5t: X5T });
  const payload = enc({
    aud: `https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`,
    iss: CLIENT_ID, sub: CLIENT_ID,
    jti: randomBytes(16).toString('hex'),
    nbf: now, exp: now + 600,
  });
  const signingInput = `${header}.${payload}`;
  const sig = createSign('RSA-SHA256').update(signingInput).sign(PEM).toString('base64url');
  const res = await fetch(`https://login.microsoftonline.com/${TENANT_ID}/oauth2/v2.0/token`, {
    method: 'POST',
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
    body: new URLSearchParams({
      grant_type: 'client_credentials',
      client_id: CLIENT_ID,
      client_assertion_type: 'urn:ietf:params:oauth:client-assertion-type:jwt-bearer',
      client_assertion: `${signingInput}.${sig}`,
      scope: `${SPO_ORIGIN}/.default`,
    }).toString(),
  });
  const j = await res.json();
  if (!j.access_token) throw new Error('token: ' + JSON.stringify(j));
  return j.access_token;
}

const TOKEN = await appOnlyToken();
const auth = { Authorization: `Bearer ${TOKEN}`, Accept: 'application/json;odata=nometadata' };

// ---------- 3. PDF generators -------------------------------------------

async function buildPdf({ title, lines, links = [], creationSeed = 0 }) {
  const pdf = await PDFDocument.create({ updateMetadata: false });
  // pdf-lib stamps creationDate/modificationDate from the system clock by
  // default — disable that so the PDF is byte-stable across runs (we want
  // to deliberately produce DIFFERENT bytes for v1 vs v2 only when the
  // content differs).
  pdf.setCreationDate(new Date('2020-01-01T00:00:00Z'));
  pdf.setModificationDate(new Date('2020-01-01T00:00:00Z'));
  const page = pdf.addPage([612, 792]);
  const font = await pdf.embedFont(StandardFonts.Helvetica);
  const fontBold = await pdf.embedFont(StandardFonts.HelveticaBold);

  let y = 740;
  page.drawText(title, { x: 50, y, font: fontBold, size: 18, color: rgb(0, 0, 0) });
  y -= 30;
  for (const line of lines) {
    page.drawText(line, { x: 50, y, font, size: 11, color: rgb(0, 0, 0), maxWidth: 500 });
    y -= 18;
  }
  y -= 10;
  for (const [display, url] of links) {
    page.drawText(display, { x: 50, y, font, size: 11, color: rgb(0, 0.4, 0.8) });
    const w = font.widthOfTextAtSize(display, 11);
    const annot = pdf.context.obj({
      Type: 'Annot', Subtype: 'Link',
      Rect: [50, y - 2, 50 + w, y + 11],
      Border: [0, 0, 0],
      A: { Type: 'Action', S: 'URI', URI: pdf.context.obj(url) },
    });
    const ref = pdf.context.register(annot);
    const existing = page.node.Annots() ?? pdf.context.obj([]);
    existing.push(ref);
    page.node.set(pdf.context.obj('Annots'), existing);
    y -= 22;
  }
  return Buffer.from(await pdf.save({ useObjectStreams: false, addDefaultPage: false }));
}

// ---------- 4. SP file upload --------------------------------------------

function spSafe(part) {
  // Per-segment URL encoding; double single quotes for OData literals.
  return encodeURIComponent(part).replace(/'/g, "''");
}

async function uploadFile(sitePath, libraryServerRel, fileName, bytes) {
  const folderEncoded = libraryServerRel
    .split('/').filter(Boolean).map(spSafe).join('/');
  const url =
    `${SPO_ORIGIN}${sitePath}/_api/web/getfolderbyserverrelativeurl('/${folderEncoded}')` +
    `/files/add(url='${spSafe(fileName)}',overwrite=true)`;
  const res = await fetch(url, {
    method: 'POST',
    headers: { ...auth, 'Content-Type': 'application/octet-stream' },
    body: bytes,
  });
  if (!res.ok) {
    throw new Error(`upload ${sitePath}/${libraryServerRel}/${fileName}: ${res.status} ${(await res.text()).slice(0, 300)}`);
  }
  const j = await res.json();
  return { serverRelativeUrl: j.ServerRelativeUrl, length: j.Length };
}

// ---------- 5. SP sharing-link creation ---------------------------------

// SP REST endpoint that returns an org-view sharing link (`:b:/s/<site>/<token>`)
// for an absolute file URL. Tested working with app-only Sites.FullControl.All.
//
// Anonymous-link endpoints (SP.Web.CreateAnonymousLink) are blocked by
// most tenant policies — we don't try them. Org-view is the shape most
// SPO admins lose sleep over because the token is opaque, the link can
// outlive the file, and it doesn't show up in the file's Manage Access
// UI.
async function createOrgSharingLink(sitePath, absoluteFileUrl) {
  const res = await fetch(
    `${SPO_ORIGIN}${sitePath}/_api/SP.Web.CreateOrganizationSharingLink`,
    {
      method: 'POST',
      headers: { ...auth, 'Content-Type': 'application/json;odata=nometadata' },
      body: JSON.stringify({ url: absoluteFileUrl, isEditLink: false }),
    }
  );
  if (!res.ok) {
    throw new Error(`createSharingLink ${absoluteFileUrl}: ${res.status} ${(await res.text()).slice(0, 300)}`);
  }
  const j = await res.json();
  return j.value;
}

// Enumerate files in a site's Documents library so we don't have to
// hardcode names that drift over time.
async function listLibraryFiles(sitePath, libraryFolderRel = 'Shared Documents') {
  const folder = `${sitePath}/${libraryFolderRel}`.split('/').filter(Boolean).map(spSafe).join('/');
  const r = await fetch(
    `${SPO_ORIGIN}${sitePath}/_api/web/getfolderbyserverrelativeurl('/${folder}')/files?$select=Name,ServerRelativeUrl,TimeLastModified`,
    { headers: auth },
  );
  if (!r.ok) return [];
  return (await r.json()).value;
}

// ---------- 6. Stale-duplicate flow -------------------------------------

console.log('==> Generating vendor-onboarding.pdf v1 and v2');
const v1 = await buildPdf({
  title: 'Vendor Onboarding Procedures',
  lines: [
    'Last updated: 2019. Yes, really. Janet has been meaning to update.',
    '',
    'Step 1: New vendor fills out form W-9 and emails it to ap@contoso.com.',
    'Step 2: Finance mails back the confirmation packet (in 4-6 weeks).',
    'Step 3: Vendor mails back the signed confirmation (in another 4-6 weeks).',
    'Step 4: Wait for someone to file the paperwork. They will eventually.',
    '',
    'Note: We do not currently have an online portal. Yes, we know.',
    'The "expense report deadline" is 5 PM Friday. This is firm.',
  ],
  links: [],
});
const v2 = await buildPdf({
  title: 'Vendor Onboarding Procedures',
  lines: [
    'Last updated: 2024. We finally got around to it.',
    '',
    'Step 1: New vendor self-registers at vendors.contoso.com/register',
    'Step 2: AP team reviews within 2 business days.',
    'Step 3: Approval email triggers automated DocuSign for W-9.',
    'Step 4: System notifies requester. End-to-end: ~3 days.',
    '',
    'IMPORTANT: The expense report deadline changed in 2023 to NOON Friday.',
    'If you are still hitting the 5 PM target, please update your bookmarks.',
    'Several departments did not get the memo. We are working on it.',
  ],
  links: [],
});
console.log(`    v1 size: ${v1.length} bytes`);
console.log(`    v2 size: ${v2.length} bytes`);

console.log('\n==> Uploading v1 to /sites/demo-hr (this will be the STALE copy)');
const hrStale = await uploadFile('/sites/demo-hr', '/sites/demo-hr/Shared Documents', 'vendor-onboarding.pdf', v1);
console.log(`    ${hrStale.serverRelativeUrl} (${hrStale.length} bytes)`);

console.log('\n==> Uploading v1 to /sites/demo-finance (will be overwritten with v2 next)');
const financeV1 = await uploadFile('/sites/demo-finance', '/sites/demo-finance/Shared Documents', 'vendor-onboarding.pdf', v1);
console.log(`    ${financeV1.serverRelativeUrl} (${financeV1.length} bytes)`);

// Sleep briefly so SP records v1 as a distinct version before we overwrite.
await new Promise((r) => setTimeout(r, 1500));

console.log('\n==> Overwriting finance copy with v2 (creates version history)');
const financeV2 = await uploadFile('/sites/demo-finance', '/sites/demo-finance/Shared Documents', 'vendor-onboarding.pdf', v2);
console.log(`    ${financeV2.serverRelativeUrl} (${financeV2.length} bytes — version 2)`);
console.log('    Once the duplicates bootstrap runs, demo-hr/vendor-onboarding.pdf');
console.log('    should appear in Stale because its current hash matches finance\'s prior version.');

// ---------- 7. Sharing-link flow -----------------------------------------

console.log('\n==> Generating org-view sharing links for existing demo files');

// Walk the Documents library of each demo site and grab the first PDF
// we see — that's the file we'll mint a sharing link for. This is more
// robust than hardcoded paths because the demo-fixtures folder names
// drift over time as we tweak content.
const DEMO_SITES = [
  { site: '/sites/demo-engineering', label: 'Engineering' },
  { site: '/sites/demo-hr', label: 'HR' },
  { site: '/sites/demo-marketing', label: 'Marketing' },
  { site: '/sites/demo-finance', label: 'Finance' },
];

const sharingLinks = [];
for (const s of DEMO_SITES) {
  const files = await listLibraryFiles(s.site, 'Shared Documents');
  // Skip the v1/v2 stale-copy files since we just touched those.
  const candidates = files.filter((f) => f.Name.endsWith('.pdf') && !f.Name.includes('vendor-onboarding'));
  if (candidates.length === 0) {
    console.log(`    ${s.label}: no PDF in Shared Documents to share, skipping`);
    continue;
  }
  const target = candidates[0];
  const absoluteUrl = `${SPO_ORIGIN}${target.ServerRelativeUrl}`;
  try {
    const link = await createOrgSharingLink(s.site, absoluteUrl);
    console.log(`    ${s.label} (${target.Name}): ${link}`);
    sharingLinks.push({
      site: s.site,
      label: `${s.label} — ${target.Name.replace(/\.pdf$/i, '')}`,
      path: target.ServerRelativeUrl,
      sharingUrl: link,
    });
  } catch (err) {
    console.log(`    ${s.label} (${target.Name}): ${err.message.split('\n')[0]}`);
  }
}

if (sharingLinks.length === 0) {
  console.log('\n    No sharing links generated. Check that the target files exist and the');
  console.log('    Entra app has Sites.FullControl.All (Application) consented.');
  process.exit(0);
}

// ---------- 8. Build a new page that embeds the sharing-link URLs --------

console.log('\n==> Creating /sites/demo-engineering/SitePages/Cross-Site-References.aspx');

const pageBody = `
<h2>Cross-Site References</h2>
<p>The links below are a mix of <em>direct</em> AllItems-style URLs and SharePoint
<em>sharing wrappers</em> (the <code>:b:/s/...</code> shape that comes out of the
"Share" → "People in your organization with the link" button). Same files,
different URL shapes — useful for seeing what the inventory finds and what it
classifies as "sharing wrapper".</p>

<h3>Sharing-wrapper links (the ones SPO admins lose sleep over)</h3>
<ul>
${sharingLinks.map((l) => `<li><a href="${l.sharingUrl}">${l.label} (sharing link)</a></li>`).join('\n')}
</ul>

<h3>Same files, direct links</h3>
<ul>
${sharingLinks.map((l) => {
  const allItems = `${SPO_ORIGIN}${l.site}/Shared%20Documents/Forms/AllItems.aspx?id=${encodeURIComponent(l.path)}&parent=${encodeURIComponent(l.path.split('/').slice(0, -1).join('/'))}`;
  return `<li><a href="${allItems}">${l.label} (direct)</a></li>`;
}).join('\n')}
</ul>

<h3>Cross-site stale-copy demo</h3>
<p>The HR site has an old copy of <a href="${SPO_ORIGIN}/sites/demo-hr/Shared%20Documents/Forms/AllItems.aspx?id=/sites/demo-hr/Shared%20Documents/vendor-onboarding.pdf&amp;parent=/sites/demo-hr/Shared%20Documents">Vendor Onboarding Procedures</a>
that hasn't been updated since 2019. Finance has the
<a href="${SPO_ORIGIN}/sites/demo-finance/Shared%20Documents/Forms/AllItems.aspx?id=/sites/demo-finance/Shared%20Documents/vendor-onboarding.pdf&amp;parent=/sites/demo-finance/Shared%20Documents">current 2024 version</a>,
but the HR copy is what shows up first in search. Janet says she'll get to it.</p>
`.trim();

// Use the SitePages REST API: addsitepage → savepageasdraft → publish.
async function addSitePage(sitePath, fileName, title, body) {
  // Check if the page already exists; if so, skip the create step.
  const existsRes = await fetch(
    `${SPO_ORIGIN}${sitePath}/_api/sitepages/pages/getbyurl(url='SitePages/${spSafe(fileName)}')`,
    { headers: auth },
  );
  let pageId;
  if (existsRes.status === 404) {
    const createRes = await fetch(
      `${SPO_ORIGIN}${sitePath}/_api/sitepages/pages`,
      {
        method: 'POST',
        headers: { ...auth, 'Content-Type': 'application/json;odata=nometadata' },
        body: JSON.stringify({
          PageLayoutType: 'Article',
          Title: title,
          Name: fileName,
        }),
      }
    );
    if (!createRes.ok) {
      throw new Error(`addsitepage failed: ${createRes.status} ${(await createRes.text()).slice(0, 200)}`);
    }
    pageId = (await createRes.json()).Id;
  } else if (existsRes.ok) {
    pageId = (await existsRes.json()).Id;
  } else {
    throw new Error(`page-exists check failed: ${existsRes.status}`);
  }

  // Build canvas content: a single text web part containing the HTML body.
  const canvas = [
    {
      controlType: 4,
      position: { layoutIndex: 1, zoneIndex: 1, sectionIndex: 1, sectionFactor: 12, controlIndex: 1 },
      innerHTML: body,
    },
  ];

  await fetch(
    `${SPO_ORIGIN}${sitePath}/_api/sitepages/pages/getbyurl(url='SitePages/${spSafe(fileName)}')/checkoutpage`,
    { method: 'POST', headers: auth },
  );
  const saveRes = await fetch(
    `${SPO_ORIGIN}${sitePath}/_api/sitepages/pages/getbyurl(url='SitePages/${spSafe(fileName)}')/savepageasdraft`,
    {
      method: 'POST',
      headers: { ...auth, 'Content-Type': 'application/json;odata=nometadata' },
      body: JSON.stringify({
        Title: title,
        CanvasContent1: JSON.stringify(canvas),
      }),
    }
  );
  if (!saveRes.ok) throw new Error(`savepageasdraft: ${saveRes.status} ${(await saveRes.text()).slice(0, 300)}`);
  const pubRes = await fetch(
    `${SPO_ORIGIN}${sitePath}/_api/sitepages/pages/getbyurl(url='SitePages/${spSafe(fileName)}')/publish`,
    { method: 'POST', headers: auth },
  );
  if (!pubRes.ok) throw new Error(`publish: ${pubRes.status} ${(await pubRes.text()).slice(0, 300)}`);
  return pageId;
}

const pageId = await addSitePage('/sites/demo-engineering', 'Cross-Site-References.aspx', 'Cross-Site References', pageBody);
console.log(`    page id ${pageId} published`);

console.log('\n✅ Seed complete.');
console.log('   Run the duplicates bootstrap on /sites/demo-finance Documents lib');
console.log('   (Duplicates tab → Bootstrap version history) to surface the stale pair.');
console.log('   Run a unified scan to ingest the new sharing-link references.');

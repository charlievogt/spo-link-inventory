import * as React from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Link } from '@fluentui/react/lib/Link';

export class HelpTab extends React.Component {
  public render(): React.ReactElement {
    return (
      <Stack tokens={{ childrenGap: 12 }} style={{ padding: 16, maxWidth: 760 }}>
        <Text variant="xLarge" style={{ fontWeight: 600 }}>About Link Inventory</Text>

        <Text>
          This web part shows two reports backed by the same Azure Function:
        </Text>

        <Stack tokens={{ childrenGap: 8 }}>
          <Text>
            <strong>Link Inventory.</strong> Walks every site you have access to,
            extracts every link from every modern page and Office/PDF document,
            and lets you search, filter, and find-and-replace across the
            entire tenant. Use it to find broken links, count references to
            a target, or migrate a URL pattern across hundreds of pages
            at once.
          </Text>
          <Text>
            <strong>Duplicates.</strong> Hashes every file in the inventoried
            libraries and groups by hash to find exact duplicates. Detects
            "stale copies" by walking version history (a previous version of
            file A matches the current content of file B → A is stale).
            Detects same-name pairs across sites for triage.
          </Text>
        </Stack>

        <Text variant="large" style={{ fontWeight: 600, marginTop: 12 }}>How scans work</Text>
        <Text>
          Scans run as background jobs in your Azure Function. Each job
          processes one site at a time off a Storage Queue, so a tenant
          scan won't time out a single function invocation. Scan results
          live in Azure Blob Storage; the web part polls a status endpoint
          while running and renders the aggregate when the job finishes.
          Daily delta scans (configured in the Schedule panel on the
          Link Inventory tab) skip any file or page whose ETag/Modified
          timestamp matches the last scan, so cost stays low after the
          initial baseline.
        </Text>

        <Text variant="large" style={{ fontWeight: 600, marginTop: 12 }}>Permissions</Text>
        <Text>
          Most users see their own readable subset of sites and can browse
          inventory + duplicate reports for those sites. Members of the
          configured admin Entra group can additionally trigger scans,
          run find-and-replace, manage the duplicates allowlist, and
          delete jobs.
        </Text>

        <Text variant="large" style={{ fontWeight: 600, marginTop: 12 }}>Project</Text>
        <Text>
          Source, issues, deploy guide:{' '}
          <Link href="https://github.com/charlie-vogt/spo-link-inventory" target="_blank" rel="noopener noreferrer" data-interception="off">
            github.com/charlie-vogt/spo-link-inventory
          </Link>
        </Text>
      </Stack>
    );
  }
}

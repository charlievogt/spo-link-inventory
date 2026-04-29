// Test setup loaded via `node --import ./dist-test/src/__tests__/test-setup.js`
// before any test module evaluates. Provides default env values that
// some services read at module-initialization time (e.g. backlinksIndex).
process.env.SPO_TENANT_HOST ??= "contoso.sharepoint.com";
process.env.LINK_INVENTORY_ADMIN_GROUP_ID ??= "00000000-0000-0000-0000-000000000000";

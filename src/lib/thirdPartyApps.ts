// Well-known third-party (non-Microsoft) applications that turn up in
// business email compromise investigations: mail clients with full mailbox
// sync, backup/export tools and address-book harvesters. Not Microsoft apps:
// shown as third-party, never with the Microsoft badge.
// Source: https://byteintocyber.com/microsoft-365-application-ids-bec-investigation-resources/
// Retrieved 2026-11-08.

export interface ThirdPartyApp {
  appId: string;
  displayName: string;
  // Why it matters in an investigation.
  note: string;
}

export const THIRD_PARTY_APPS: ThirdPartyApp[] = [
  {
    appId: 'a43e5392-f48b-46a4-a0f1-098b5eeb4757',
    displayName: 'CloudSponge',
    note: 'Address book exfiltration',
  },
  {
    appId: '62db40a4-2c7e-4373-a609-eda138798962',
    displayName: 'Edison Mail',
    note: 'Email client with full synchronisation capabilities. UAL typically shows items synced identified by ‘MailItemsAccessed’ events.',
  },
  {
    appId: 'e9a7fea1-1cc0-4cd9-a31b-9137ca5deedd',
    displayName: 'eM Client',
    note: 'Email client with full synchronisation capabilities. UAL typically shows items synced identified by ‘MailItemsAccessed’ events. (Search App ID)',
  },
  {
    appId: '77468577-4f6e-40e7-b745-11d3d0c28095',
    displayName: 'Fastmail',
    note: 'Mailbox exfiltration/persistence',
  },
  {
    appId: '2ef68ccc-8a4d-42ff-ae88-2d7bb89ad139',
    displayName: 'Mail_Backup',
    note: 'Updated version of PERFECTDATA SOFTWARE. See that row for additional information.',
  },
  {
    appId: 'a245e8c0-b53c-4b67-9b45-751d1dff8e6b',
    displayName: 'Newsletter Software Supermailer',
    note: 'Bulk email sending software.',
  },
  {
    appId: 'ff8d92dc-3d82-41d6-bcbd-b9174d163620',
    displayName: 'PERFECTDATA SOFTWARE',
    note: 'Backup/Export mailboxes. UAL does not show items synced. UPDATE: It now shows items synced. Identified by MailItemsAccessed events. (Search App ID)',
  },
  {
    appId: '179d5108-412b-4c95-8e34-06786784ab39',
    displayName: 'PostBox',
    note: 'Email client with full synchronisation capabilities.',
  },
  {
    appId: '4761b959-9780-4c2d-87a3-512b4638f767',
    displayName: 'Rclone',
    note: 'Manage files within M365.',
  },
  {
    appId: 'caffae8c-0882-4c81-9a27-d1803af53a40',
    displayName: 'SigParser',
    note: 'Address book exfiltration',
  },
  {
    appId: '946c777c-bc85-489e-b034-392389ae23d6',
    displayName: 'Spike',
    note: 'Mailbox exfiltration/persistence',
  },
  {
    appId: '497ac034-5120-4c1a-929a-0351f5c09918',
    displayName: 'ZoomInfo Communitiez Login',
    note: 'Address book exfiltration',
  },
  {
    appId: '858d7e42-35f0-44b7-9033-df309239a47f',
    displayName: 'Zoominfo Login',
    note: 'Address book exfiltration',
  },
];

const byId = new Map(THIRD_PARTY_APPS.map((a) => [a.appId.toLowerCase(), a]));

export function getThirdPartyApp(appId: string): ThirdPartyApp | undefined {
  return byId.get(appId.trim().toLowerCase());
}

// SAMPLE DATA ONLY — made-up projects and emails for the browser preview.
// In Outlook, projects will come from the SharePoint Project Register and the
// filed-email list once the app registration is set up.

window.SampleData = {
  projects: [
    {
      hdNumber: 'HD4203', jobName: '23 Ellis Street, Frankton DD', client: 'Clientco Ltd',
      status: 'Site Investigation', clientContacts: ['sarah@clientco.co.nz'],
      siteUrl: 'https://hdhq.sharepoint.com/sites/HD4203', folders: ['Email'],
    },
    {
      hdNumber: 'HD4141', jobName: 'Clientco subdivision stage 2', client: 'Clientco Ltd',
      status: 'Assessment / Reporting', clientContacts: ['james@clientco.co.nz'],
      siteUrl: 'https://hdhq.sharepoint.com/sites/HD4141', folders: ['Email'],
    },
    {
      hdNumber: 'HD4141-4', jobName: 'Additional testing (variation)', client: 'Clientco Ltd',
      status: 'New', clientContacts: ['james@clientco.co.nz'], motherHd: 'HD4141',
      siteUrl: 'https://hdhq.sharepoint.com/sites/HD4141', folders: ['HD4141-4 Variation/Email', 'Email'],
    },
    {
      hdNumber: 'HD4188', jobName: '12 River Road, Hamilton East', client: 'Waikato Builders Ltd',
      status: 'Site Investigation', clientContacts: ['tom@waikatobuilders.co.nz'],
      siteUrl: 'https://hdhq.sharepoint.com/sites/HD4188', folders: ['Email'],
    },
    {
      hdNumber: 'HD3990', jobName: 'Te Rapa warehouse', client: 'Waikato Builders Ltd',
      status: 'Completed', clientContacts: ['tom@waikatobuilders.co.nz'],
      siteUrl: 'https://hdhq.sharepoint.com/sites/HD3990', folders: ['Email'],
    },
  ],

  // Emails already filed (this becomes the "Project Email Register").
  filedLog: [
    {
      messageId: 'm-001', conversationId: 'conv-ellis', hdNumber: 'HD4203', folder: 'Email',
      subject: 'HD4203 – retaining wall levels', from: 'neil@hdgeo.co.nz', fromName: 'Neil Alsop',
      client: 'Clientco Ltd', date: '2026-11-18',
      preview: 'Hi Sarah, can you send through the latest surveyed levels for the retaining wall?',
    },
    {
      messageId: 'm-002', conversationId: 'conv-setout', hdNumber: 'HD4203', folder: 'Email',
      subject: 'Wall setout queries', from: 'sarah@clientco.co.nz', fromName: 'Sarah Ngata',
      client: 'Clientco Ltd', date: '2026-10-02',
      preview: 'A couple of questions on the setout of the boundary wall before the contractor starts.',
    },
    {
      messageId: 'm-004', conversationId: 'conv-levels', hdNumber: 'HD4203', folder: 'Email',
      subject: 'Surveyed levels for the retaining wall', from: 'sarah@clientco.co.nz', fromName: 'Sarah Ngata',
      client: 'Clientco Ltd', date: '2026-10-21',
      preview: 'Surveyor has picked up the existing ground levels along the back boundary. Spreadsheet attached.',
    },
    {
      messageId: 'm-005', conversationId: 'conv-cpt', hdNumber: 'HD4141-4', folder: 'HD4141-4 Variation/Email',
      subject: 'Extra CPTs – variation approved', from: 'james@clientco.co.nz', fromName: 'James Tane',
      client: 'Clientco Ltd', date: '2026-11-19',
      preview: 'Approved. Please go ahead with the two extra CPTs near lot 14.',
    },
    {
      messageId: 'm-006', conversationId: 'conv-gate', hdNumber: 'HD4188', folder: 'Email',
      subject: 'Gate code and site access', from: 'tom@waikatobuilders.co.nz', fromName: 'Tom Walker',
      client: 'Waikato Builders Ltd', date: '2026-11-04',
      preview: 'Gate code is 4821. Please text me before you arrive so I can move the digger.',
    },
    {
      messageId: 'm-003', conversationId: 'conv-river', hdNumber: 'HD4188', folder: 'Email',
      subject: 'HD4188 borehole logs', from: 'rhys@hdgeo.co.nz', fromName: 'Rhys Christian',
      client: 'Waikato Builders Ltd', date: '2026-11-10',
      preview: 'Attached are the draft borehole logs from Tuesday. Groundwater was met at 2.4 m.',
    },
  ],

  // Emails to try in the preview, one per situation the add-in must handle.
  emails: [
    {
      label: 'HD number in subject',
      messageId: 'e-1', conversationId: 'conv-a', itemClass: 'IPM.Note',
      subject: 'HD4188 – updated site plan', from: 'tom@waikatobuilders.co.nz', fromName: 'Tom Walker',
      to: ['rhys@hdgeo.co.nz'], date: '2026-11-23',
      body: 'Hi Rhys, the architect has moved the garage. Updated site plan attached.',
    },
    {
      label: 'Client reply, no HD number (same thread)',
      messageId: 'e-2', conversationId: 'conv-ellis', itemClass: 'IPM.Note',
      subject: 'RE: retaining wall levels', from: 'sarah@clientco.co.nz', fromName: 'Sarah Ngata',
      to: ['neil@hdgeo.co.nz'], date: '2026-11-23',
      body: 'Hi Neil, attached are the updated levels from the surveyor. Can you confirm the wall design still works?',
    },
    {
      label: 'New email from a client with one active job',
      messageId: 'e-3', conversationId: 'conv-b', itemClass: 'IPM.Note',
      subject: 'Site access next week', from: 'tom@waikatobuilders.co.nz', fromName: 'Tom Walker',
      to: ['rhys@hdgeo.co.nz'], date: '2026-11-23',
      body: 'Gate code is changing on Monday. New code to follow.',
    },
    {
      label: 'Client with several active jobs',
      messageId: 'e-4', conversationId: 'conv-c', itemClass: 'IPM.Note',
      subject: 'Question about your report', from: 'mike@clientco.co.nz', fromName: 'Mike Brown',
      to: ['neil@hdgeo.co.nz'], date: '2026-11-23',
      body: 'Could you explain the bearing capacity section in a bit more detail?',
    },
    {
      label: 'Dash job (variation)',
      messageId: 'e-5', conversationId: 'conv-d', itemClass: 'IPM.Note',
      subject: 'HD4141-4 additional testing scope', from: 'james@clientco.co.nz', fromName: 'James Tane',
      to: ['ben@hdgeo.co.nz'], date: '2026-11-23',
      body: 'Please go ahead with the extra CPTs we discussed.',
    },
    {
      label: 'Meeting invite (should not file)',
      messageId: 'e-6', conversationId: 'conv-e', itemClass: 'IPM.Schedule.Meeting.Request',
      subject: 'HD4188 site meeting', from: 'rhys@hdgeo.co.nz', fromName: 'Rhys Christian',
      to: ['tom@waikatobuilders.co.nz'], date: '2026-11-23',
      body: 'Site meeting to walk through the investigation locations.',
    },
    {
      label: 'Unknown sender',
      messageId: 'e-7', conversationId: 'conv-f', itemClass: 'IPM.Note',
      subject: 'Quick question', from: 'someone@gmail.com', fromName: 'A. Person',
      to: ['queries@hdgeo.co.nz'], date: '2026-11-23',
      body: 'Do you do geotech reports for house sites in Cambridge?',
    },
  ],
};

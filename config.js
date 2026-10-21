/* ===================================================================
   Rincovitch Site Inspections — settings for your Microsoft 365
   Your IT administrator fills in the four values below once.
   See "SETUP GUIDE - IT.txt" for where each value comes from.
   =================================================================== */
window.APP_CONFIG = {

  // 1. Directory (tenant) ID of your Microsoft 365 organisation
  tenantId: 'PASTE-TENANT-ID-HERE',

  // 2. Application (client) ID of the app registration made for this app
  clientId: 'PASTE-CLIENT-ID-HERE',

  // 3. Sharing link to the SharePoint / OneDrive FOLDER that holds the app data and reports
  //    (Share > Copy link on the folder, for "People in Rincovitch" or specific people with edit access)
  dataFolderLink: 'PASTE-FOLDER-LINK-HERE',

  // 4. Sharing link to the inspection register workbook (.xlsx) — rows are added and updated live.
  //    Leave as '' to skip live register updates.
  registerLink: '',

  // FOR A ONE-PERSON TEST ON YOUR OWN ONEDRIVE you can use folder/file paths instead of links
  // (then leave dataFolderLink / registerLink as they are). Paths are inside "My files", e.g.:
  //   dataFolderPath: 'Inspections App',
  //   registerPath:   'Inspections App/MEL03-Inspection-Register.xlsx',
  dataFolderPath: '',
  registerPath: '',

  // Optional
  reportsFolder: 'Inspections',          // sub-folder (inside the data folder) where each inspection's Report/ and Photos/ are filed
  registerTable: 'InspectionRecords',    // name of the table in the register workbook
  registerSheet: 'Inspection Register',  // name of the sheet holding that table
  pollSeconds: 45,                        // how often to check for other inspectors' changes while the app is open
};

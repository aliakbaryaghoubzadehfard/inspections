/* ===================================================================
   Rincovitch Site Inspections — settings for your Microsoft 365
   Your IT administrator fills in the four values below once.
   See "SETUP GUIDE - IT.txt" for where each value comes from.
   =================================================================== */
window.APP_CONFIG = {

  // 1. Directory (tenant) ID of your Microsoft 365 organisation
  tenantId: 'consumers',                 // personal Microsoft accounts

  // 2. Application (client) ID of the app registration made for this app
  clientId: 'c1fcd549-0956-4b5c-83d7-d849a62eb4e2',

  // 3. Sharing link to the SharePoint / OneDrive FOLDER that holds the app data and reports
  //    (Share > Copy link on the folder, for "People in Rincovitch" or specific people with edit access)
  dataFolderLink: 'PASTE-YOUR-INSPECTIONS-APP-FOLDER-LINK-HERE',

  // 4. The inspection register workbook (.xlsx) — rows are added and updated live.
  //    Simplest: keep the workbook INSIDE the data folder and give its file name here (no second link needed):
  registerInFolder: 'MEL03-Inspection-Register.xlsx',
  //    Or, if the workbook lives somewhere else, paste its sharing link here instead (and set registerInFolder to ''):
  registerLink: '',

  // FOR A ONE-PERSON TEST ON YOUR OWN ONEDRIVE you can use folder/file paths instead of links
  // (then leave dataFolderLink / registerLink as they are). Paths are inside "My files", e.g.:
  //   dataFolderPath: 'Inspections App',
  //   registerPath:   'Inspections App/MEL03-Inspection-Register.xlsx',
  dataFolderPath: 'Inspections App',     // used only until the folder link above is filled in (then the link wins)
  registerPath: '',

  // Optional
  reportsFolder: 'Inspections',          // sub-folder (inside the data folder) where each inspection's Report/ and Photos/ are filed
  registerTable: 'InspectionRecords',    // name of the table in the register workbook
  registerSheet: 'Inspection Register',  // name of the sheet holding that table
  pollSeconds: 30,                        // how often to check for other inspectors' changes while the app is open
};

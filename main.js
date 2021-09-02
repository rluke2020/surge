// for excel
var XLSX = require('xlsx')

const { BrowserWindow, Menu, app, shell, dialog, ipcMain } = require('electron')

// MENU TEMPLATE
let template = [{
  label: 'Links',
  click: () => {
    app.emit('activate')
  }
}]



var fs = require('fs')
// Modules to control application life and create native browser window
const { spawn } = require('child_process')



const path = require('path')


function createWindow() {
  // Create the browser window.
  const mainWindow = new BrowserWindow({
    width: 800,
    height: 600,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      nodeIntergration: true
    }
  })

  // and load the index.html of the app.
  mainWindow.loadFile('index.html')

  // Open the DevTools.
  // mainWindow.webContents.openDevTools()
}

// This method will be called when Electron has finished
// initialization and is ready to create browser windows.
// Some APIs can only be used after this event occurs.
app.whenReady().then(() => {
  createWindow()

  app.on('activate', function () {
    // On macOS it's common to re-create a window in the app when the
    // dock icon is clicked and there are no other windows open.
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

// Quit when all windows are closed, except on macOS. There, it's common
// for applications and their menu bar to stay active until the user quits
// explicitly with Cmd + Q.
app.on('window-all-closed', function () {
  if (process.platform !== 'darwin') app.quit()
})


// WHEN APP IS READY
app.on('ready', () => {
  // const menu = Menu.buildFromTemplate(template)
  // Menu.setApplicationMenu(menu)
})

// In this file you can include the rest of your app's specific main process
// code. You can also put them in separate files and require them here.

// save dialog 
destination()

extract()


open_file_dialog()

filesFolder()

consolidate()

function consolidate() {
  ipcMain.on('consolidate', (event, dir, files, destination, startDate, sheetToConsolidate) => {

    let arg = dir

    // collumn to search for a date
    let dateCollumn

    loadEvent(event)

    let array_of_sheets = ["Active Index Testing ", "OPD Screening", "Missed Appointments ", "Defaulter Q1"]

    let finalsheet = []

    let wb = XLSX.utils.book_new()
    wb.Props = {
      Title: "surge cascade ",
      Subject: "Output",
      Author: "UI-HEXED",
      CreatedDate: new Date(),
    }

    array_of_sheets.forEach(element => {

      sheetToConsolidate = element
      finalsheet = [];
      finalsheet.push(  heads( sheetToConsolidate ) )
      // console.log(  heads( sheetToConsolidate ) );

      files.forEach((element, index) => {


        let path = `${arg}\\${element}`

        if (sheetToConsolidate === 'Defaulter Q1') {
          dateCollumn = 'F'
        } else {
          dateCollumn = 'E'
        }

        var sheet = XLSX.readFile(path, { sheets: sheetToConsolidate })


        let data = sheet.Sheets[sheetToConsolidate]


        XLSX.utils.sheet_to_json(data)

        // parsing date to be compatible with workboook dates
        let check = Date.parse(startDate[sheetToConsolidate]).toString().substr(0, 5)
        let check2 = parseInt(check) + 6
        let check2string = check2.toString()


        let taken = []


        // this will check every key / or cell in the sheeet
        for (const key in data) {

          if (key[0] == dateCollumn && (Date.parse(data[key].w).toString().substr(0, 5) == check || Date.parse(data[key].w).toString().substr(0, 5) == check2string)) {
            let row = key.substr(1, 100)


            for (const key in data) {

              if (row == key.substr(1, 100)) {
                // this if statement will check date collumn to output correct date

                taken.push(data[key].w)


              }

            }

          }
          if (taken.length > 0) {
            finalsheet.push(taken)
            taken = []
          }
        }

      })

      let sheet_of_all_files = XLSX.utils.aoa_to_sheet(finalsheet)
      sheet_of_all_files['!rows'] = [{hpx:40 }]
      sheet_of_all_files['!cols'] = [{width:40 },{width:40 },{width:40 },{width:40 },{width:40 },{width:40 },{width:40 },{width:40 },{width:40 },{width:40 },{width:40 },{width:40 },{width:40 },{width:40 },{width:40 },{width:40 },{width:40 }]
      

      finalsheet = []

      XLSX.utils.book_append_sheet(wb, sheet_of_all_files, sheetToConsolidate)

    })

    XLSX.write(wb, { bookType: 'xlsx', type: 'binary' })

    XLSX.writeFile(wb, destination)

    event.sender.send('result', 'finished consolidating')


  })
}

function filesFolder() {
  ipcMain.on('filesFolder', (event) => {
    let result = dialog.showOpenDialog({
      properties: ['openDirectory']
    })

    result.then((value) => {

      if (!value.canceled) {

        loadEvent(event)




        //variable to contain files in the directory
        let newFiles = []


        fs.readdir(value.filePaths[0], function (err, files) {

          if (err) {
            // do stuff to handle the error
            return event.sender.send('result', 'unable to do necessary stuff : ' + err)

          }
          let patt = new RegExp("^[A-z].*.xlsx")



          files.forEach(file => {

            const filtered = patt.exec(file)


            if (filtered) {
              newFiles.push(filtered[0])
            }

          })
          event.sender.send('selected-consolidate', value.filePaths, newFiles)
          event.sender.send('result', `Found ${newFiles.length} files`)
        })

      }

    })
  })
}

function open_file_dialog() {
  ipcMain.on('open-file-dialog', (event) => {


    let result = dialog.showOpenDialog({
      properties: ['openFile']
    })
    result.then((value) => {

      if (!value.canceled) {

        loadEvent(event)

        var workbook = XLSX.readFile(value.filePaths[0])
        var sheet_name_list = workbook.SheetNames
        event.sender.send('sheets', sheet_name_list)

        event.sender.send('selected-directory', value.filePaths)
      }

    })

  })
}

function extract() {
  ipcMain.on('extract', (event, arg) => {

    // first read fie
    let sheet = XLSX.readFile(arg[0],  { sheets: [arg[2]] , cellDates: true })


    // creating a new workbok
    var wb = XLSX.utils.book_new()

    // get the workbook data
    let data = sheet.Sheets[arg[2]]

    XLSX.utils.book_append_sheet(wb, data, arg[2])

    // write to new file
    XLSX.write(wb, { bookType: 'xlsx', type: 'binary', cellDates:true })

    XLSX.writeFile(wb, `${arg[1]}`)

    event.sender.send('result', 'finished Extracting sheet')


  })
}

/**
 * @param void
 * 
 * @returns null
 * 
 * @
 * 
 * function for getting destination
 */
function destination() {
  ipcMain.on('destination', (event) => {
    let result = dialog.showOpenDialog({
      properties: ['openDirectory']
    })
    event.sender.send('result', 'selecting directory...')

    result.then((value) => {
      loadEvent(event)

      if (!value.canceled) {

        event.sender.send('selected-destination', value.filePaths)
      }

    })
  })
}

function loadEvent(event) {

  event.sender.send('loading')
}

function heads(sheet) {
  let headers = {
    "Active Index Testing ": [
      "District", 
      "Site", "District_site",
      "UID",
      "Cohort Week Start",
      "Cohort Week End",
      "Cohort graduation date  (28-day follow-up)",
      "Reporting Due Date/ Biweekly report due",
      "Surge Call date",
      "# newly diagnosed positives in the specified cohort week",
      "# clients already on ART  offered AIT in the cohort week",
      "Total offered AIT -- # new positives and clients already on ART offered AIT in the cohort week",
      "Accepted index testing-- Of those offered, # who accepted index testing",
      "% accepted index testing",
      "# Contacts Elicited -- Of those who accepted index testing, total contacts identified",
      "Contact elicitation ratio",
      "# Contacts eligible for index testing -- Of contacts listed, total  reported to be HIV negative or unknown status",
      "Contacts reached-- # of contacts that were contacted and reached",
      "Contacts eligible for testing -- HIV status Ascertainment",
      "Contacts Tested -- # of eligible contacts who received HIV testing services",
      "% Tested -- Percent of contacts who received HIV testing services",
      "Tested Positive -- # contacts who received  HTS and tested positive",
      "% Yield",
      "Linked to ART- Among new positives, # linked to ART",
      "I. % Linked(Target: ≥ 95%)",
      "Describe specific problems & gaps related to index testing at this site",
      "Remediation Plan",
      "Responsible person(s)",
      "Timeframe for remediating problem",
      "Current status/Update"
    ],
    "OPD Screening": [
      "District",
      "Site", "District_site",
      "UID",
      "Week Start",
      "Week End",
      "Report due date. Same as Bi-weekly reporting date",
      "OPD attendance - # of clients that attended  the OPD clinic",
      "Clients Screened- of clients that attended OPD, # screened for HIV testing eligibility",
      "% Screened",
      "Eligible for testing- Of those screened, # eligible for testing",
      "% Eligible for Testing",
      "Tested- Of those eligible, # of clients who accepted testing",
      "% tested",
      "Tested Positive- # who received HTS and tested positive",
      "Yield (Target: ≥ 10%)",
      "Linked to ART- Among new positives, # linked to ART ",
      "I. % Linked(Target: ≥ 95%)",
      "Describe specific problems & gaps related to OPD screening at this site",
      "Remediation Plan",
      "Responsible person(s)",
      "Timeframe for remediating problem",
      "Current status/Update"
    ],
    "Missed Appointments ": [
      "District",
      "Site", "District_site",
      "UID",
      "Cohort Week Start",
      "Cohort Week End",
      "Cohort graduation date  (28-day follow-up)",
      "Reporting Due Date/ Biweekly report due ",
      "Surge Call date ",
      "# of patients who missed appointment (cohort denominator)",
      "# patients who missed appiontment ≥ 14 days (tracing denominator)",
      "No contact tracing attempted - # who missed appointment ≥ 14 days and  not traced ",
      "Of patients not traced, # with no physical address or phone number",
      "# patients who missed appointments traced by phone only",
      "# patients who missed appointments traced physically only - home visit ",
      "# patients who missed appointment traced by both phone and home visit",
      "Total Attempted to Trace",
      "Successfully traced",
      "% traced",
      "Back to Care -- Traced and brought back to care",
      "Self Return to Care -- Not traced but returned to care",
      "% Back in Care",
      "Self-Transferred Out",
      "Died",
      "Stopped ART",
      "Promised to return but not yet visited clinic",
      "On ART, due to ART gap (poor adherence)",
      "On ART, no gap (e.g., received emergency supply)",
      "LTFU-- Total number not reachable (by phone, home visit, or both)",
      "% LTFU",
      "Describe specific problems & gaps related to missed appt tracing at this site",
      "Remediation Plan",
      "Responsible person(s)",
      "Timeframe for remediating problem",
      "Current status/Update"
    ],
    "Defaulter Q1":["District ", "Site ", "UID", "Current Reporting Quarter", "Data Collection Date", "Reporting Due Date", "Previous Reporting Quarter", "Total TX_CURR- previous reporting quarter", "Generate PEPFAR TX_CURR report from EMR and enter total ", "Total Defaulters- Generate the EMR defaulter list and enter total number", "False Defaulters - of total defaulters # with an appointment that had not been entered in EMR)", "Alive in Care", "T/O", "Died", "Stop", "Duplicate", "True Defaulters- of total defaulters # truly 2 months late for scheduled appt", "No contact tracing attempted ", "% Not Contacted", "Of those not traced, # with no phone number or physical address in record", "# defaulters traced by phone only", "# defaulters traced physically only - home visit ", "# defaulters traced by both phone and home visit", "Total Attempted to Trace", "Successfully traced", "% traced", "Back to Care -- Traced and brought back to care", "Self Return to Care -- Not traced but returned to care", "% Back in Care", "Self-Transferred Out", "Died", "Stopped ART", "Promised to return but not yet visited clinic", "On ART due to ART gap (poor adherence)", "On ART, received emergency supply", "LTFU- Total not reachable (by phone, home visit, or both)", "% LTFU", "Describe specific problems & gaps related to missed appt tracing at this site", "Remediation Plan", "Responsible person(s)", "Timeframe for remediating problem", "Current status/Update"
    ]
  };

  return headers[sheet]

}
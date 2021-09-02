// All of the Node.js APIs are available in the preload process.
// It has the same sandbox as a Chrome extension.
var { ipcRenderer } = require('electron')

window.addEventListener('DOMContentLoaded', () => {
  const replaceText = (selector, text) => {
    const element = document.getElementById(selector)
    if (element) element.innerText = text
  }

  for (const type of ['chrome', 'node', 'electron']) {
    replaceText(`${type}-version`, process.versions[type])
  }

  document.getElementsByTagName('input').


  Selected_file = 'path'
  // for open file
  const selectDirBtn = document.getElementById('select-directory')

  // for files
  const selectfilesBtn = document.getElementById('select-files-directory')

  // for extracting
  const extract = document.getElementById('extract')

  //for opening destination

  const destination = document.getElementById('destination_btn')

  // options

  // request destination folder

  try {
    destination.addEventListener('click', (event) => {
      ipcRenderer.send('destination')
    })


  } catch (error) {
    // console.error(error);
  }

  //  resuest files folder
  try {

    selectfilesBtn.addEventListener('click', (event) => {
      ipcRenderer.send('filesFolder');
    })

  } catch (error) {
    // console.error(error);
  }

  // request extract action
  try {
    extract.addEventListener('click', (event) => {
      let source = document.getElementById('selected-file').innerText

      let destination = document.getElementById('destination').innerText

      let sheet_index = document.getElementById('select').value

      let start_date =  '1'  //document.getElementById('start').value

      let min_row =  '1'  //document.getElementById('min-row').value
      let max_row =  '1'  //document.getElementById('max-row').value
      let max_column =  '1'  //document.getElementById('max-column').value
      let min_column =  '1'  //document.getElementById('min-column').value
      let loading = document.getElementById('loading');

      loading.innerHTML = 'trying to extract...'

      // return 1;
      ipcRenderer.send('extract', [source, destination, sheet_index, start_date, min_row, max_row, max_column, min_column])

    })
  } catch (error) {
    //  console.error(error);
  }


  try {
    selectDirBtn.addEventListener('click', (event) => {
      ipcRenderer.send('open-file-dialog')
      ipcRenderer.send('asynchronous-message')
    })
  } catch (error) {
    // console.error(error);
  }


  // when destination folder is selected
  ipcRenderer.on('selected-destination', (event, path) => {
    let destination = document.getElementById('destination')
    let rename = document.getElementById('rename')
    let loading = document.getElementById('loading');

    loading.innerHTML = ''
    destination.innerText = path.toString()
    destination.innerText += "\\" + rename.value
  })

  ipcRenderer.on('selected-consolidate', (event, dir, files) => {
    let consolidate = document.getElementById('consolidate')
  
    if (files.length > 0) {
      consolidate.disabled = false

    // get start date
    //  let start = document.getElementById('start')
    //  console.log( start.value );

    //  get destination folder
    let destination = document.getElementById('destination')




      consolidate.addEventListener( 'click' , (event)=>{
        let dates = document.getElementsByTagName('input')
        let datesObject = {
    
        };
    
        for (let index = 0; index < dates.length; index++) {
           datesObject[dates[index].name] = dates[index].value;
          
        }
    
        console.log(datesObject);
        
        // get sheet name
        let sheet = document.getElementById('select').value
        


        ipcRenderer.send('consolidate', dir, files , destination.innerText , datesObject, sheet)

      } )
    }


  })

  // when the directory is selected
  ipcRenderer.on('selected-directory', (event, path) => {

    let selected_file = document.getElementById('selected-file')
    selected_file.innerText = `${path}`

    let file = selected_file.innerText.split("\\")
    let rename = document.getElementById('rename')

    rename.value = file[file.length - 1]

  })

  // when the sheets have been retrieved
  ipcRenderer.on('sheets', (event, data) => {
    // console.log(data);
    let list = data.toString().split(',') //split the sheets into various sheets
    let sheets = document.getElementById('select');
    let x = 0
    let loading = document.getElementById('loading');

    loading.innerHTML = 'extracting'

    // for each sheet add to the drop down list
    list.forEach(item => {
      sheets.innerHTML += (`<option value='${item}'>${item}</option>`)

      x++
    });
    loading.innerHTML = ''
  })


  ipcRenderer.on('asynchronous-reply', (event, arg) => {
    const message = `Asynchronous message reply: ${arg}`
    console.log(message);
  })

  ipcRenderer.on('loading', (event, arg) => {
    let loading = document.getElementById('loading');

    loading.innerHTML = 'please a wait moment...'

  })

  // message to get when the extracting job is finished
  ipcRenderer.on('result', (event, arg) => {
    let loading = document.getElementById('loading');

    loading.innerHTML = arg
  })




})






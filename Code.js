function onOpen() {
 SpreadsheetApp
   .getUi()
   .createMenu("Library")
   .addItem("Index Library", "indexLibrary")
   .addItem("Propose Titels", "showIncompleteTitels")
   .addItem("Fix Titels", "fixTitels")
   .addToUi();
}


function indexLibrary(){
  const all_folders = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('folders').getDataRange().getValues();
  const books_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('books');
  let books = books_sheet.getDataRange().getValues();

  let completed_folders = [];
  books.forEach(b => {if (b[0]) completed_folders.push(b[0])} );

  let folders = [];
  all_folders.forEach(f => { 
    if(!completed_folders.includes(f[0]))
      folders.push(f);
  });

  const chunks = chunk(folders, 100);  
  for(let i=0; i < chunks.length; i++){
    for (let n=0; n < chunks[i].length; n++){
      let bsif = DriveApp.getFolderById(chunks[i][n][0]).getFiles();
      while(bsif.hasNext()) {
        let bid = bsif.next().getId();
        let idx = books.findIndex(b => b[2] == bid);
        if (idx){
          books[idx][0] = chunks[i][n][0];
          books[idx][1] = chunks[i][n][1];
        } else {
          console.log(`Could not find book with id ${bid} in books sheet data`);
        }
        
      };
    }
    books_sheet.getDataRange().setValues(books);
  }
}

function showIncompleteTitels(){
  let books = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('books').getDataRange().getValues();

  books.shift();
  if(!books[0][0]){
    SpreadsheetApp.getActiveSpreadsheet().toast('Before trying to generate title proposals, please first index the library', 'Notice');
    return;
  }

  books = books.filter((b) => (b[4] == 'no'));
  let updates = books.map(function(b){
    let filename = b[3],
      author = b[1],
      title = filename.replace(/\.[\w]{3,4}/ig, '') //remove file ending
      .replace(/\[.*?\]|\(.*?\)/ig, '') //remove anything in between brackets or parenthesis
      .replace(/\bby\b.*/iu, ' ') //remove anything after 'by'
      .replace(/[\p{P}\s]+/giu, ' ') //remove punctuation and extra whitespcae
      .replace(new RegExp(author.replace(/[\p{P}\s]+/giu, ' '),'gi'), ' ') //remove author name
      .trim();

    return [b[2], b[3], b[1], title, '', ''];
  });

  const updates_sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('updates');

  //reset sheet for new data
  updates_sheet.clear();
  updates_sheet.setFrozenRows(1);
  updates_sheet.getRange(1,1,1,6).setValues([['Book ID',	'Book Name',	'Author Name',	'Title',	'Year',	'Ready?']]).setFontWeight("bold");
  updates_sheet.getRange(2,1,updates.length,updates[0].length).setValues(updates);
}

function fixTitels(){
  let updates = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('updates').getDataRange().getValues();

  if(updates.length < 2){
    SpreadsheetApp.getActiveSpreadsheet().toast('Please use "Propose Titles" first and then mark titles to be renamed by adding "ok" in the column "Ready?.', 'Notice');
  }

  updates.shift();
  updates = updates.filter( b => (b[5] == 'ok') );

  if(updates.length < 1){
    SpreadsheetApp.getActiveSpreadsheet().toast('Please mark titles to be renamed first by adding "ok" in the column "Ready?".', 'Notice');
    return;
  }

  updates.forEach(function(up){
    let file = DriveApp.getFileById(up[0]);
    let file_ending = file.getMimeType().match(/pdf/gi) ? '.pdf' : '.epub';
    let fn_parts = [up[3], "["+up[2].trim()+"]"];
    if (up[4]) {
      fn_parts.push("("+(up[4]+'').trim()+")");
    }
    //console.log('old '+up[1]+' new: '+fn_parts.join('_')+file_ending);
    file.setName(fn_parts.join('_')+file_ending);
  });
}

function chunk(arr, chunkSize) {
  if (chunkSize <= 0) throw "Invalid chunk size";
  let chunks = [];
  for (let i=0; i<arr.length; i+=chunkSize)
    chunks.push(arr.slice(i,i+chunkSize));
  return chunks;
}
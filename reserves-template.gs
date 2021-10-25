//Run automatically on open
function triggerOnOpen(e) {
  main();
}

//this determines which column is edited, and can call other functions if conditions are met

// this fires on any cell edit, and is connected to the trigger
function triggerOnEdit(e)
{
  switchboard(e);
}



function main() {
  //Load the columns we want to read and edit
  var sheet = SpreadsheetApp.getActive();
  var titleColumn = sheet.getRange('E3:E');
  var authorColumn = sheet.getRange('F3:F');
  var editionColumn = sheet.getRange('G3:G');
  var isbnColumn = sheet.getRange('D3:D');
  var queryColumn = sheet.getRange('T3:T');
  var notesColumn = sheet.getRange('K3:K');
  var printCopyColumn = sheet.getRange('I3:I');
  var eAccessColumn = sheet.getRange('J3:J');
  var purchasePrintColumn = sheet.getRange('L3:L');
  var purchaseEbookColumn = sheet.getRange('M3:M');
  
  var blankCount = 0;
  
  //Loop through the isbn column
  for (let i=1; i<=isbnColumn.getValues().length; i+=1) {
    var isbn = String(isbnColumn.getCell(i,1).getValue());
    //Check if a query has been completed for the isbn
    if (!isbnColumn.getCell(i,1).isBlank() && queryColumn.getCell(i,1).isBlank()) {
      var isbn = String(isbnColumn.getCell(i,1).getValue());
      Logger.log(isbn);
      
      //Get alternate isbns from worldcat
      var notFound = null;
      var worldCatData = parseWorldCat(isbn);
      if (titleColumn.getCell(i,1).isBlank()) {
        titleColumn.getCell(i,1).setValue(worldCatData['title']);
      }
      if (authorColumn.getCell(i,1).isBlank()) {
        authorColumn.getCell(i,1).setValue(worldCatData['author']);
      }
      var isbns = worldCatData['isbns'];
      if (isbns[0]=='ISBN not found in WorldCat') {
        notFound= isbns[0];
        isbns = Array(isbns[1]);
        //notesColumn.getCell(f,1).setValue(isbns);
        //queryColumn.getCell(f,1).setValue('done');
        //continue;
      }
      //Search primo for our inventory
      var inLibrary = parseSRULoop(isbns);
      Logger.log(inLibrary);
      queryColumn.getCell(i,1).setValue('done');
      let notes = inLibrary[0];
      if (notFound != null) {
        notes += '   |  ' + notFound;
      }
      for (let j = 1; j<inLibrary.length; j+=1) {
        if (j==1){
          notes += '  |  Alt ISBNs: '+inLibrary[j]
        }
        else {
          notes+= ', '+inLibrary[j];
        }
      }
      notesColumn.getCell(i,1).setValue(notes);
      if (worldCatData['title']!=undefined && titleColumn.getCell(i,1).isBlank()) {
        titleColumn.getCell(i,1).setValue(worldCatData['title']);
      }
      if (worldCatData['author']!=undefined && authorColumn.getCell(i,1).isBlank()) {
        titleColumn.getCell(i,1).setValue(worldCatData['author']);
      }
      
      //Write result for print books
      if (inLibrary.some(item => item.includes('Physical owned'))) {
        printCopyColumn.getCell(i,1).setValue('Owned');
        purchasePrintColumn.getCell(i,1).setValue('No action required');
      }
      else {
        printCopyColumn.getCell(i,1).setValue('Not Owned');
      }

      //Write result for ebooks
      if (inLibrary.some(item => item.includes('E-Access available'))) {
        if (inLibrary.some(item => item.includes('Unlimited'))) {
          eAccessColumn.getCell(i,1).setValue('Unlimited Users');
          purchaseEbookColumn.getCell(i,1).setValue('No action required');
        }
        else if (inLibrary.some(item => item.includes('Limited to 3'))) {
          eAccessColumn.getCell(i,1).setValue('3 Users');
        }
        else if (inLibrary.some(item => item.includes('Limited to 1'))) {
          eAccessColumn.getCell(i,1).setValue('Single User');
        }
        else if (inLibrary.some(item => item.includes('DDA Pool'))) {
          eAccessColumn.getCell(i,1).setValue('DDA Pool');
        }
        else if (inLibrary.some(item => item.includes('Unknown'))) {
          eAccessColumn.getCell(i,1).setValue('Unknown');
        }
      }
      else {
        eAccessColumn.getCell(i,1).setValue('No access');
      }
    }
    else if (isbnColumn.getCell(i,1).isBlank()) {
      blankCount+=1;
    }
    if (blankCount>=5) {
      break;
    }
  }

}






//this determines which column is edited, and can call other functions if conditions are met
function switchboard(e)
{
  var range = e.range;
  var edited_column = range.getColumn();  
  var val=range.getValue();
  var edited_row = range.getRow();
  
  if(edited_column==12 && (val=="Purchase" || val=="Purchase additional copy")){
    sendPurchaseEmail(e, edited_row);
  }
  
  //more conditions could be added below, that point to other/new functions
  
 if(edited_column==17){
    sendEbookActivationEmail(e, edited_row);
  }

}


function sendPurchaseEmail(e, edited_row){
  
  var emailAddress="jjacobs@lclark.edu";
  var subject="Purchase Request for Course Reserves";
  
  var rowData = SpreadsheetApp.getActiveSheet().getRange(edited_row, 1, 1, 11).getValues();
  var course=rowData[0][0];
  var instructor=rowData[0][2];
  var title=rowData[0][4];
  var isbn=rowData[0][3];
  var edition=rowData[0][7];
  var notes=rowData[0][11];
  
  var message="Hello! \n\n";
  message += "Please purchase the following item for course reserves:\n";
  message += "Title: "+title+"\n";
  message += "ISBN: "+isbn+"\n\n";
    message += "edition: "+edition+"\n\n";
  message += "Please flag this for the following course:\n";
  message += "Course: "+course+"\n";
  message += "Instructor: "+instructor+"\n\n";
  message += "Thank you!"
  
  MailApp.sendEmail(emailAddress,subject,message);


}

function sendEbookActivationEmail(e, edited_row){

var emailAddress="jjacobs@lclark.edu";
  var subject="Ebook Activated for Course Reserves";
  
  var rowData = SpreadsheetApp.getActiveSheet().getRange(edited_row, 1, 1, 5).getValues();
  var course=rowData[0][0];
  var instructor=rowData[0][2];
  var title=rowData[0][3];
  var isbn=rowData[0][4];
  
  var message="Hello! \n\n";
  message += "An ebook has been activated for course reserves:\n";
  message += "Title: "+title+"\n";
  message += "ISBN: "+isbn+"\n\n";
  message += "Please add it to the following reading list:\n";
  message += "Course: "+course+"\n";
  message += "Instructor: "+instructor+"\n\n";
  message += "Thank you!"
  
  MailApp.sendEmail(emailAddress,subject,message);
  
}





function parseWorldCat(origisbn = '9780811819046') {
  //Use given isbn to gather other possible isbns from worldcat
  //worldcay query
  var url = 'http://www.worldcat.org/webservices/catalog/search/worldcat/opensearch?q='+origisbn+'&wskey=OOoG4G5MohZ2DjaaeeNBeSUkWhYwa54i9I7EqbRTxkCoWucaqCRSe1Sj1mKDdcRy2QY6jrWv2bGywXTP';
  //get worldcat response
  var xmlWorldCat = UrlFetchApp.fetch(url).getContentText();
  
  // get isbn info from xml response
  var doc = XmlService.parse(xmlWorldCat);
  var root = doc.getRootElement();
  var rootnamespace = root.getNamespace();
  var entries = root.getChildren('entry', rootnamespace);
  var content = root.getAllContent();
  var worldCatData = {};
  if (content.length==0) {
    worldCatData['isbns'] = Array('ISBN not found in WorldCat', origisbn);
    return worldCatData;
  }
  var isbns = null;
  var title=null;
  var author=null;
  for (let i=0 ; i < entries.length; i +=1) {
    let entry = entries[i];
    let allentry = entry.getChildren();
    var isbnnamespace = null;
    for (let j=0 ; j < allentry.length; j+=1) {
      let child = allentry[j];
      if (child.getName() == 'identifier') {
        isbnnamespace = child.getNamespace();
        break;
      }
    }
    if (isbnnamespace!=null) {
      var isbnsxml = entry.getChildren('identifier', isbnnamespace);
      isbns = Array();
      for (let k=0; k<isbnsxml.length; k+=1) {
        let isbn = isbnsxml[k].getText().split(':').slice(-1)[0];
        isbns.push(isbn);
      }
      if (title==null) {
        title = entry.getChild('title', rootnamespace).getText();
        Logger.log(title);
      }
      if (author==null) {
        author = entry.getChild('author', rootnamespace).getChild('name', rootnamespace).getText();
      }
      
    }
  }
  if (isbns==null) {
    isbns = Array(origisbn);
  }
  isbns = isbns.filter(item => item !== origisbn);
  isbns.unshift(origisbn);
  worldCatData['isbns'] = isbns;
  worldCatData['title'] = title;
  worldCatData['author'] = author;
  return worldCatData;
}

function parseSRU(isbn = '9780811819046') {
  //query primo for our inventory
  var url = 'https://na01.alma.exlibrisgroup.com/view/sru/01ALLIANCE_LCC?version=1.2&operation=searchRetrieve&query=alma.isbn='+ isbn;
  
  //read xml response and get information on matrial type
  var xmlWorldCat = UrlFetchApp.fetch(url).getContentText();
  var doc = XmlService.parse(xmlWorldCat);
  var root = doc.getRootElement();
  var rootnamespace = root.getNamespace();
  var numberOfRecords = root.getChild('numberOfRecords', rootnamespace).getText();
  if (numberOfRecords == 0) {
    return Array('not owned');
  }
  var recordData = root.getChild('records',rootnamespace).getChild('record', rootnamespace).getChild('recordData', rootnamespace);
  var recordIdentifier = root.getChild('records',rootnamespace).getChild('record', rootnamespace).getChild('recordIdentifier', rootnamespace);
  var recordnamespace = recordData.getChildren()[0].getNamespace();
  var records = recordData.getChildren('record', recordnamespace);
  var results = Array();
  for (let i=0; i<records.length; i+=1) {
    var record = records[i]; 
    var datafields = record.getChildren('datafield', recordnamespace);
    for (let j=0; j<datafields.length; j+=1) {
      let datafield = datafields[j];
      if (datafield.getAttribute('tag').getValue() == 'AVA') {
        results.push('Physical owned');
      }
      else if (datafield.getAttribute('tag').getValue() == 'AVE') {
        var eAccessType = 'Unknown';
        var subfields = datafield.getChildren('subfield', recordnamespace);
        for (let k=0; k<subfields.length; k+=1) {
          if (subfields[k].getAttribute('code').getValue()=='m'){
            var eBookType = subfields[k].getText();
            if (eBookType =='JSTOR eBook Collection (Watzek purchased)' || eBookType=='ProQuest eBook Collection (Purchase on Demand)') {
              eAccessType = 'DDA Pool';
            }
            else {
              var mmsid = recordIdentifier.getText();
              var jsonurl = 'https://watzek.lclark.edu/reservesLG/amApi.php?mmsid='+mmsid;
              var jsonResponse = JSON.parse(UrlFetchApp.fetch(jsonurl).getContentText());
              var access = String(jsonResponse[mmsid]);
              Logger.log(access);
              if (access != 'false') {
                eAccessType = access;
              }
            }
          }
        }
        results.push('E-Access available'+ ' (' + eAccessType +')');      
      }
    }
    if (results.length==0){
      results.push('Owned (type not listed)');
    }
  }
  return results;
}


function parseSRULoop(isbns) {
  //query primo for each possible isbn
  var primoRecords = Array();
  for (let i=0; i < isbns.length; i+=1) {
    var isbn = isbns[i];
    var sruRecords = parseSRU(isbn);
    for (let j=0; j<sruRecords.length; j+=1) {
      primoRecords.push(isbn + ' ' +sruRecords[j]);
    }
  }
  return primoRecords;
}








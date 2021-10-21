//Run automatically on open
function triggerOnOpen(e) {
  main();
}

function main() {
  //Load the columns we want to read and edit
  var sheet = SpreadsheetApp.getActive();
  var titleColumn = sheet.getRange('D3:D');
  var authorColumn = sheet.getRange('E3:E');
  var editionColumn = sheet.getRange('F3:F');
  var isbnColumn = sheet.getRange('G3:G');
  var queryColumn = sheet.getRange('R3:R');
  var notesColumn = sheet.getRange('J3:J');
  var printCopyColumn = sheet.getRange('H3:H');
  var eAccessColumn = sheet.getRange('I3:I');
  var purchasePrintColumn = sheet.getRange('K3:K');
  var purchaseEbookColumn = sheet.getRange('L3:L');
  
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
/*function switchboard(e)
{
  var range = e.range;
  var edited_column = range.getColumn();  
  var val=range.getValue();
  var edited_row = range.getRow();
  
  if(edited_column==9 && val=="Purchase"){
    sendPurchaseEmail(e, edited_row);
  }
  
  //more conditions could be added below, that point to other/new functions
  

}
*/

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







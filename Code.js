var ANALYZER_PREFIX_SHEETDOCS = "ANALYZER_SHEETDOCS_";
var ANALYZER_CODES_STORAGE = "ANALYZER_STORAGEDOC";
var ANALYZER_PREFIX_CODE = "#CODE_";

/**
 * Creates a custom menu in Google Sheets when the spreadsheet opens.
 */
function onOpen() {
  if(DocumentApp.getActiveDocument()) {
    //TESTDOC
    PropertiesService.getDocumentProperties().setProperty(ANALYZER_CODES_STORAGE,"1E0uCRw7ga_7jF1xrclfWiQHpLAJ1CmM6z77CozfoSzU");
    DocumentApp.getUi().createMenu('#Analyzer')
      .addItem('OPEN', 'openCoder')
      .addItem('MARK', 'mark')
      .addToUi();
  } else {
    SpreadsheetApp.getUi().createMenu('#Analyzer')
      .addItem('Manage GDoc', 'showPicker')
      .addItem('Count #tags', 'showTagCountSidebar')
      .addItem('Filter by #tags', 'filterByTag')
      .addItem('Edit #tags', 'editTags')
      .addToUi();
  }
}

/**
 * Displays an HTML-service dialog in Google Sheets that contains client-side
 * JavaScript code for the Google Picker API.
 */
function showPicker() {
  var html = HtmlService.createTemplateFromFile('Picker2.html')
      .evaluate()
      .setWidth(600)
      .setHeight(425)
      .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, 'Manage Sheet\'s GDocs');
}

/**
 * Gets the user's OAuth 2.0 access token so that it can be passed to Picker.
 * This technique keeps Picker from needing to show its own authorization
 * dialog, but is only possible if the OAuth scope that Picker needs is
 * available in Apps Script. In this case, the function includes an unused call
 * to a DriveApp method to ensure that Apps Script requests access to all files
 * in the user's Drive.
 *
 * @return {string} The user's OAuth 2.0 access token.
 */
function getOAuthToken() {
  DriveApp.getRootFolder();
  return ScriptApp.getOAuthToken();
}


function processDocs() {
  var docs = getActiveSheetDocs();
  Logger.log("PROCESS:" + JSON.stringify(docs));
  if(docs){
    retrieveComments(Object.keys(docs));
  } else {
    sheet.getRange(1, 1).setValue('NO DOC WAS SELECTED! NO DATA, NO FUN! =/');
  }
}


function getActiveSheetDocs() {
  var documentProperties = PropertiesService.getDocumentProperties();
  var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var propID = ANALYZER_PREFIX_SHEETDOCS + activeSheet.getSheetId();
  var docsList = documentProperties.getProperty(propID);

  //Logger.log("ACTIVE_SHEET_DOCS: " + docsList);
  //documentProperties.deleteProperty(propID);
  //return {};
  return JSON.parse(docsList);
}


function addDocToSheet(docId,docName) {
  if(docId) {
    var documentProperties = PropertiesService.getDocumentProperties();
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var propID = ANALYZER_PREFIX_SHEETDOCS + activeSheet.getSheetId();
    var docs = documentProperties.getProperty(propID);
    if (!docs) {
      docs = {};
    } else {
      docs = JSON.parse(docs);
    }
    if (!(docId in docs)) {
      docs[docId] = docName;
    }
    documentProperties.setProperty(propID, JSON.stringify(docs));
  }
}


function removeDocFromSheet(docID) {
  Logger.log("REMOVE_DOC: "+ docID);
  if(docID) {
    var documentProperties = PropertiesService.getDocumentProperties();
    var activeSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var propID = ANALYZER_PREFIX_SHEETDOCS + activeSheet.getSheetId();
    var docs = documentProperties.getProperty(propID);
    if (docs) {
      docs = JSON.parse(docs);
      delete docs[docID];
      documentProperties.setProperty(propID, JSON.stringify(docs));
    }
  }
}


function showTagCountSidebar() {
  var html = HtmlService.createTemplateFromFile('TagCount')
      .evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle('Tag Count')
      .setWidth(430)
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Retrieve a list of comments.
 *
 * @param {Array} docIDs List of document IDs to retrieve comments from.
 */
function retrieveComments(docIDs) {
  var info = [];
  var callArguments = {'maxResults': 100, 'fields': 'items(commentId,content,context/value,fileId,replies/content,status),nextPageToken'}
  var docComments, pageToken;
  for (var doc=0; doc < docIDs.length; doc++) {
    do {
      callArguments['pageToken'] = pageToken;
      docComments = Drive.Comments.list(docIDs[doc],callArguments);
      info = info.concat(getCommentsInfo(docIDs[doc],docComments.items));
      pageToken = docComments.nextPageToken;
    } while(pageToken);
  }    
  writeComments(info);
}

function writeComments(comments) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  sheet.clear();
  for(var line=0; line < comments.length; line++) {
    sheet.getRange(line+1, 1, 1, comments[line].length).setValues([comments[line]]);
  }
}


function getCommentsInfo(fileId,commentsList) {
  var result = []
  var docURL = DriveApp.getFileById(fileId).getUrl();
  for(var line=0; line < commentsList.length; line++) {
    var comment = JSON.parse(commentsList[line]);
    if(comment['status']==='open') {
      var tags = getTags(comment);
      if(tags.length>0) {
        result.push(tags);
        var link = docURL + '&disco='+comment['commentId'];
        result.push(['=HYPERLINK(\"'+link+'\",\"'+format(comment['context']['value'])+'\")']);
      }
    }
  }
  return result;
}


function getTags(comment){
  var tags = getHashTags(comment['content']);
  var replies = comment['replies'];
  if(replies) {
    for(var reply=0; reply<replies.length;reply++) {
      var replyTags = replies[reply]['content'];
      tags = tags.concat(getHashTags(replyTags));
    }
  }
  return tags;
}


function getHashTags(comment){
  if(typeof comment === 'string' || comment instanceof String) {
    var tags = comment.match(/#\S+/g);
    if(tags) return tags;
  }
  return [];
}

function format(text) {
  return text;
}

function escapeHTML(unsafe) {
  return unsafe
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&#039;/g, "\'");
}


function getTagCount() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var tagCount = {};
  
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  for(var row=1; row <= lastRow; row++) {
    for(var col=1; col <= lastCol; col++) {
      var tag = sheet.getRange(row,col).getValue();
      if(tag[0] !== '#') { break; }
      
      if(tagCount[tag]) {
        tagCount[tag] = tagCount[tag]+1;
      } else {
        tagCount[tag] = 1;
      }
    }
  }
  return sortTagCount(tagCount);
}


function sortTagCount(tagCount) {
  var arr = [];
  var prop;
  for (prop in tagCount) {
    if (tagCount.hasOwnProperty(prop)) {
      arr.push({
        'key': prop,
        'value': tagCount[prop]
      });
    }
  }
  arr.sort(function(a, b) {
    return b.value - a.value;
  });
  return arr; // returns array
}


function filterByTag() {
  var tags = SpreadsheetApp.getUi().prompt("Enter #tags to filter content:").getResponseText().trim();
  if(tags.length>0) {
    var filter = tags.split(" ");
    var aSheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var values = aSheet.getDataRange().getValues();
    
    var linesToCopy = [];
    for(var row = 0; row < values.length; row++) {
      if(values[row][0][0] !== '#') {continue;}
      
      var filter_tmp = filter.slice();
      
      for(var col = 0; col < values[row].length && filter_tmp.length > 0; col++) {
        
        for(var f = 0; f < filter_tmp.length; f++) {
          if(values[row][col].indexOf(filter_tmp[f]) > -1) {
            filter_tmp.splice(f,1);
            break;
          }
        }
      }
      if(filter_tmp.length==0) {
        linesToCopy.push(row+1,row+2);
      }
    }
    if(linesToCopy.length>0) {
      var sheetName = aSheet.getName()+ " - [" + filter + "]";
      var newSheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(sheetName);
      copyLinesBetweenSheets(aSheet, newSheet, linesToCopy);
    } else {
      SpreadsheetApp.getUi().alert("There were no matches for #tags: " + filter);
    }
  } else {
    SpreadsheetApp.getUi().alert("Are you kidding? =/");
  }
}


function copyLinesBetweenSheets(fromSheet,toSheet,lines) {
  var lastColumn = fromSheet.getLastColumn();
  for(var line=1; line <= lines.length; line++) {
    var origin = fromSheet.getRange(lines[line-1],1,1,lastColumn);
    if(line%2==1){ var values = origin.getValues();}
    else {var values = origin.getFormulas();}
    toSheet.getRange(line,1,1,origin.getNumColumns()).setValues(values);
  }
}


function editTags() {
  var selectedCell = SpreadsheetApp.getActiveSheet().getActiveCell();
  Logger.log("CELL: [" + selectedCell.getRow() + "," + selectedCell.getColumn() + "]");
  var ids = getTagDocAndCommentId(selectedCell.getRow(), selectedCell.getColumn());
  if(Object.keys(ids).length > 0) {
    var comment = getComment(ids['docID'],ids['comID']);
    if(comment['author']['isAuthenticatedUser']) {
      var message = "[Observe that we'll also change the original comment!]";
      var save_as = "EDIT";
    } else {
      var message = "[You cannot edit the original comment so we will add typed #tags as a reply!]";
      var save_as = "REPLY";
    }
    showEditTagsTab(message, selectedCell.getValue(), comment['content'], ids['docID'], ids['comID'], save_as);
  } else {
    SpreadsheetApp.getUi().alert("We failed to find what to edit: \nPlease, select a spreadsheet cell with a tag you want to edit!");
  }
}


function getTagDocAndCommentId(cellRow, cellColumn) {
  selectedCell = SpreadsheetApp.getActiveSheet().getRange(cellRow, cellColumn, 1);
  if(selectedCell.getValue().indexOf("#") != 0){
    return {};
  } else {
    firstCellNextRow = SpreadsheetApp.getActiveSheet().getRange(cellRow+1, 1, 1);
    var cellFormula = firstCellNextRow.getFormula();
  }

  if(cellFormula.indexOf("=HYPERLINK") > -1){
    var url = cellFormula.split("\"")[1].split("/");
    var docID = url[5];
    var comID = url[6].split("=")[2];
    return {docID : docID, comID : comID};
  } else {
    return {};
  }
}


function getComment(docID, comID) {
  callArguments = {'fields': 'author/isAuthenticatedUser,commentId,content,replies(author/isAuthenticatedUser,replyId,content)'};
  return Drive.Comments.get(docID, comID,callArguments);
}


function showEditTagsTab(message, selectedTag, commentTags, doc_id, com_id, save_as) {
  var html = HtmlService.createTemplateFromFile('EditTags');
  html.message = message;
  html.tag = selectedTag;
  html.tags = commentTags.trim();
  html.doc_id = doc_id;
  html.com_id = com_id;
  html.submit_as = save_as;
  html = html.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(400);
  SpreadsheetApp.getUi().showModalDialog(html, "Edit tags...");
}


function saveTagEdition(tags, doc_id, com_id, reply_id, save_as) {
  Logger.log("SAVE: " + tags + " - " + save_as);
  var result;
  var comment = {'content':tags};
  if(save_as === 'EDIT'){
    if(reply_id) {
      result = Drive.Replies.update(comment, doc_id, com_id, reply_id);
    } else {
      result = Drive.Comments.update(comment, doc_id, com_id);
    }
  } else if(save_as === 'REPLY'){
    result = Drive.Replies.insert(comment, doc_id, com_id);
  }
  return result;
}


function changeTagInstance(originalTag, newTagValue, docID, comID) {
  comment = getComment(docID, comID);
  var content = comment['content'];
  Logger.log("CHANGE: " + content);
  var canEdit = comment['author']['isAuthenticatedUser'];
  var result = false;
  if(content.indexOf(originalTag) > -1) {
    if(canEdit) {
      result = saveTagEdition(content.replace(originalTag, newTagValue), docID, comID, null, 'EDIT');
    } else {
      result = saveTagEdition("[X"+originalTag+"-->] " + newTagValue, docID, comID, null, 'REPLY');
    }
  } else {
    var replies = comment['replies'];
    var done = false;
    for(var i=0; !done && i<replies.length; i++) {
      content = replies[i]['content'];
      canEdit = replies[i]['author']['isAuthenticatedUser'];
      if(content.indexOf(originalTag) > -1) {
        if(canEdit) {
          result = saveTagEdition(content.replace(originalTag, newTagValue), docID, comID, replies[i]['replyId'], 'EDIT');
        } else {
          result = saveTagEdition("#X"+originalTag+"--> "+newTagValue, docID, comID, null, 'REPLY');
        }
        done = true;
      }
    }
  }
  return result;
}


function changeAllInstancesOfTag(originalTag, newTagValue) {
  var allData = SpreadsheetApp.getActiveSheet().getDataRange().getValues();
  for (var i = 0; i<allData.length; i=i+2){
    for (var j = 0; j<allData[i].length; j++){ 
      if(allData[i][j] === originalTag){
        ids = getTagDocAndCommentId(i+1,1);
        result = changeTagInstance(originalTag, newTagValue, ids['docID'], ids['comID']);
        if(result) {
          updatedCommentTags(ids['docID'], ids['comID'], i+1);
        }
      }
    }
  }
}


function updatedCommentTags(docID, comID, updateRow) {
  Logger.log("UPDATE: " + updateRow);
  var callArguments = {'fields': 'content,replies(content)'};
  var comment = Drive.Comments.get(docID, comID, callArguments);
  var tags = [getTags(comment)];
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if(!updateRow) {
    updateRow = sheet.getActiveCell().getRowIndex();
  }
  sheet.getRange(updateRow, 1, 1, sheet.getLastColumn()).clear();
  sheet.getRange(updateRow, 1, 1, tags[0].length).setValues(tags);
}


/***** CODER ******/


function openCoder() {
  var html = HtmlService.createTemplateFromFile('Coder');
  html = html.evaluate()
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setTitle("#ANALYZER - CODER")
      .setWidth(450);
  DocumentApp.getUi().showSidebar(html);
}


function getCodes() {
  var properties = PropertiesService.getDocumentProperties().getProperties();
  var docURL = DocumentApp.getActiveDocument().getUrl();
  var codes = {};
  Object.keys(properties).forEach(function(key){
    var keyCheck = key.split(ANALYZER_PREFIX_CODE);
    if(keyCheck[0].length==0) {
      var bookmark = DocumentApp.getActiveDocument().getBookmark(keyCheck[1]);
      if(bookmark) {
        props = JSON.parse(properties[key]);
        props['link'] = docURL+'#bookmark='+keyCheck[1];
        codes[keyCheck[1]] = props;
      } else {
        PropertiesService.getDocumentProperties().deleteProperty(key);
      }
    }
  });
  Logger.log(codes);
  return codes;
}

function mark() {
  // Display a dialog box that tells the user how many elements are included in the selection.
  var doc = DocumentApp.getActiveDocument();
  var selection = doc.getSelection();
  var selectionText = [];
  if (selection) {
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      if (element.getElement().editAsText) {
        var text = element.getElement().editAsText();
        if (element.isPartial()) {
          text.setBackgroundColor(element.getStartOffset(), element.getEndOffsetInclusive(), "#00b2ff");
          selectionText.push(text.getText().substring(element.getStartOffset(), element.getEndOffsetInclusive()+1));
        } else {
          text.setBackgroundColor("#00b2ff");
          selectionText.push(text.getText());
        }
      }
    }
    var bookmarkOffset = (elements[0].isPartial() ? elements[0].getStartOffset() : 0);
    var bookmarkPosition = doc.newPosition(elements[0].getElement(), bookmarkOffset);
    var bookmark = doc.addBookmark(bookmarkPosition);
    var coded = {
      'selection': selectionText,
      'codes': ['#ANALIZER','#TEST']
    }
    PropertiesService.getDocumentProperties().setProperty(ANALYZER_PREFIX_CODE+bookmark.getId(), JSON.stringify(coded));
    Logger.log("STORED: "+ PropertiesService.getDocumentProperties().getProperty(ANALYZER_PREFIX_CODE+bookmark.getId()));
  } else {
    DocumentApp.getUi().alert('Nothing is selected.');
  }
}


function goToBookmark(bookmarkID){
  var doc = DocumentApp.getActiveDocument();
  var bookmark = doc.getBookmark(bookmarkID);
  if(bookmark) {
    doc.setCursor(bookmark.getPosition());
  }
}
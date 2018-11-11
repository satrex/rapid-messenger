function onInstall(e){
  onOpen(e);
}

function deleteMetaData(key){
  // get all the things and delete them in one go
  var requests = ["firstFileId","secondFileId"]
   .map (function (d) {
     return {
       deleteDeveloperMetadata: {
         dataFilter:{
           developerMetadataLookup: {
           metadataKey: d
         }}
       }};
      });

  Logger.log (JSON.stringify(requests));
  if (requests.length) {
    var result = Sheets.Spreadsheets.batchUpdate({requests:requests}, ss.getId());
    Logger.log (JSON.stringify(result));
  }
}

function setMetaData(key, value){
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var finder = ss.createDeveloperMetadataFinder();
  
  ss.addDeveloperMetadata("firstFileId", "hoge");
  ss.addDeveloperMetadata("secondFileId", "funi");
  var data=finder.withKey("firstFileId").find();
  for(var i=0; i<data.length; i++){
    
    Logger.log(data[i].getValue());
  }
}

function onOpen(e) {
  Logger.log('AuthMode: ' + e.authMode);
  var menu = SpreadsheetApp.getUi().createAddonMenu();
  if(e && e.authMode == 'NONE'){
    menu.addItem('使用開始', 'askEnabled');
  } else {
    var lang = Session.getActiveUserLocale();
    menu.addItem('URL短縮', 'generateShortUrls')
      .addItem('投稿文生成', 'createFiles')
      .addItem('メール送信', 'sendMails')
      .addItem('結果をクリア', 'clearUrls')
      .addItem('設定', 'showDialog');
    var userProps = PropertiesService.getUserProperties();
    var setDefault = userProps.getProperty("willSetDefault");
    if(setDefault == 1){
      menu.addItem('初期値設定', 'defineDefaultProperties');
    }

  };
  menu.addToUi();
};

function log_WillSetDefault(){
  var userProps = PropertiesService.getUserProperties();
  var val = userProps.getProperty("willSetDefault");
  Logger.log("willSetDefault:" + val);
}

function set_WillSetDefault_True(){
  var userProps = PropertiesService.getUserProperties();
//  userProps.deleteProperty("willSetDefault");
  userProps.setProperty("willSetDefault", 1);
}

function defineDefaultProperties(){
  var props = PropertiesService.getScriptProperties();
  logProperties();
  setDefaultProperty("bitly_token", props.getProperty("bitly_token"));
  setDefaultProperty("originUrlCol", "Z");
  setDefaultProperty("newUrlCol", "AA");
  setDefaultProperty("newIdCol", "AH");
  setDefaultProperty("nicknameCol", "D");
  setDefaultProperty("mailAddressCol", "E");
  setDefaultProperty("mailTemplate", getTemplateId(/^2./));
  setDefaultProperty("templateDocId", getTemplateId(/^1./));
}

function getTemplateIdTest(){
  Logger.log('templateDocId: ' + getTemplateId(/^1./));
  Logger.log("mailTemplate: " + getTemplateId(/^2./));
}

function getTemplateId(title){
  var ss = SpreadsheetApp.getActive();
  var ssid =ss.getId();
  Logger.log("active spreadsheet id : " + ssid);
  var ssFile = DriveApp.getFileById(ssid);
  Logger.log("file got ");
  var parents = ssFile.getParents();
  Logger.log("parents hasNext? : " + parents.hasNext());

  if(!parents.hasNext()) return null;
  
  var folder = parents.next();
  Logger.log("parent folder id : " + ss.getId());
  var files = folder.getFiles();
  while(files.hasNext()){
    var file = files.next();
    var fileName = file.getName();
    if(fileName.match(title)){      
      Logger.log("file found name : " +fileName);
      return file.getId();
    }
  }
  
  return null;
}

function setDefaultProperty(key, defaultValue){
  var props = PropertiesService.getDocumentProperties();
//  if(props.
  var prop = props.getProperty(key);
  Logger.log("property got " + key + ":" + prop);
  if(prop == null || prop === "undefined" || prop.length === 0){
    props.setProperty(key, defaultValue);
    Logger.log("property set " + key + ":" + defaultValue);
  }
}

//function onOpen() {
//  var ui = SpreadsheetApp.getUi();
//  // Or DocumentApp or FormApp.
//  ui.createMenu('キャンペーン')
//      .addItem('URL短縮', 'generateShortUrls')
//      .addItem('ドキュメント生成', 'createFiles')
//      .addItem('メール送信', 'sendMails')
//      .addItem('結果をクリア', 'clearUrls')
//      .addItem('設定', 'showDialog')
//      .addToUi();
//}

//function onOpen(e) {
//  Logger.log('AuthMode: ' + e.authMode);
//  var menu = SpreadsheetApp.getUi().createAddonMenu();
//  if(e && e.authMode == 'NONE'){
//    menu.addItem('Getting Started', 'askEnabled');
//  } else {
//    var lang = Session.getActiveUserLocale();
//    var sidebar_text = lang === 'ja' ? 'サイドバーの表示' : 'Show Sidebar';
//    menu.addItem(sidebar_text, 'showSidebar');
//  };
//  menu.addToUi();
//};


function askEnabled(){
  var lang = Session.getActiveUserLocale();
  var title = 'Your Script\'s Title';
  var msg = lang === 'ja' ? '瞬速メッセンジャーが有効になりました。ブラウザを更新してください。' : 'Rapid Messenger has been enabled.';
  var ui = SpreadsheetApp.getUi();
  ui.alert(title, msg, ui.ButtonSet.OK);
};



function clearUrls(){
  var mySheet=SpreadsheetApp.getActiveSheet(); //シートを取得

//  Browser.msgBox ("rowSheet OK");
  var newUrlCol = getNewUrlCol(); //24;
  var docIdCol = getNewIdCol();
  var lastRow =mySheet.getDataRange().getLastRow(); //シートの使用範囲のうち最終行を取得
//  Browser.msgBox ("newUrlCol OK " + newUrlCol + " " + docIdCol);

  for(var i=2;i<=lastRow;i++){
    var id = getRange(mySheet, i,docIdCol).getValue(); 
    if( id.length === 0 ) continue;
    var removingDoc = DriveApp.getFileById(id);

    if(removingDoc != null || !removingDoc.isTrashed()) {
      removingDoc.setTrashed(true);
    }
//    var newDocument = DriveApp.removeFile(id);
  }
//  Browser.msgBox ("removeFile OK");

  getRange(mySheet, 2, newUrlCol).offset(0,0,lastRow).clearContent();
  getRange(mySheet, 2, docIdCol).offset(0,0,lastRow).clearContent();

}

function showDialog() {
  var html = HtmlService.createTemplateFromFile('setting.html').evaluate()
      .setWidth(400)
      .setHeight(300);
  SpreadsheetApp.getUi() // Or DocumentApp or SlidesApp or FormApp.
      .showModalDialog(html, '設定');
}

function saveSettings(e){
  var props = PropertiesService.getDocumentProperties();
  props.setProperty("bitly_token", e.bitly_token);
  props.setProperty("mailTemplate", e.mailTemplate);
  props.setProperty("templateDocId", e.templateDocId);
  props.setProperty("newUrlCol", e.newUrlCol);
  props.setProperty("newIdCol", e.newIdCol);
  props.setProperty("originUrlCol", e.originUrlCol);
  props.setProperty("nicknameCol", e.nicknameCol);
  props.setProperty("mailAddressCol", e.mailAddressCol);

  logProperties();
}

function test(){
  var psid = "";
  postMessage(psid) ;
}

//function addProperty(){
//  var key = "PAGE_ACCESS_TOKEN";
//  var value = "EAACOrw8r3ykBAKK2NHp54f92auUkZBZALiR5HaUmmnVACiX8l8eV3AnhlUETb1naLjy87ZBGXjaBOVcQPW8WZC2y8duAXdt76eBZCMWeMZAksg5vTuY1WRKqdxCxZBgfkOoQSNhCCOisjGP0uMvu4AS8KYnSiAscwz3Hd1Nw8YoDU4jy8xJ2f1N";
//  PropertiesService.getScriptProperties().setProperty(key, value);
//}

function createFiles(){  
  var newFolder = createNewFolder();
  
  generateFiles(newFolder);
}



function isAllLetter(inputtxt)
  {
   var letters = /^[A-Za-z]+$/;
   if(inputtxt.match(letters))
     {
      return true;
     }
   else
     {
     return false;
     }
  }

function isNumber(x){ 
//  Browser.msgBox ("tyoe:" + typeof(x) );

  if( typeof(x) != 'number' && typeof(x) != 'string' )
  return false;
  else 
    return (x == parseFloat(x) && isFinite(x));
}

function getRange(sheet, row, col){
  if(col == null){throw new RangeError("アドレスが不正です:" + col + row )};
//  Browser.msgBox ("getRange start");
  if(isNumber(col)){
    Logger.log(row + "," + col);
//    Browser.msgBox ("getRange isNumber=true OK");
    return sheet.getRange(row, col);
  }
//  Browser.msgBox ("getRange isNumber=false OK");
  
  if(isAllLetter(col)) {
    var a1 = col + row;
    Logger.log(a1);
//    Browser.msgBox ("getRange isAllLetter=true OK");    
    return sheet.getRange(a1);
  }
  else {throw new RangeError("アドレスが不正です:" + col + row ) };
}

function setPropertyAsTest(){
  PropertiesService.getDocumentProperties().setProperty("test1", "foo");
  PropertiesService.getDocumentProperties().setProperty("test2", "bar");
  PropertiesService.getDocumentProperties().setProperty("test3", "baz");

  logProperties();
}

function logProperties(){
  var props = PropertiesService.getDocumentProperties();
  var keys = props.getKeys();
  Logger.log("keys:" + keys.join(","));
  keys.forEach(function(key) {
    Logger.log(key + ": " + props.getProperty(key));
  });

}

function sendMails(){
  /* スプレッドシートのシートを取得と準備 */
  var mySheet=SpreadsheetApp.getActiveSheet(); //シートを取得
  var rowSheet=mySheet.getDataRange().getLastRow(); //シートの使用範囲のうち最終行を取得

  /* メールテンプレートは独立した文書 */
  // 紹介依頼メールテンプレート
  logProperties();
  var mailTemplateId= PropertiesService.getDocumentProperties().getProperty("mailTemplate");
  Logger.log("mailTmpl:" + mailTemplateId);
  var mailTemplate = DocumentApp.openById(mailTemplateId);
  var title = getMailTemplateTitle(mailTemplate);

  for(var i=2;i<=rowSheet;i++){
    var personName = getRange(mySheet, i,NAME_COL).getValue().replace(" ", "");
    if (personName.length == 0){
      break;
    }

    var docIdCol = getNewIdCol();
    var nicknameCol = getNicknameCol();
    var mailAddressCol = getMailAddressCol();
    
    var documentId = getRange(mySheet, i,docIdCol).getValue(); // ドキュメントID 
    var fileUrl = DriveApp.getFileById(documentId).getUrl();
    var hisName = getRange(mySheet,i,nicknameCol).getValue();　//メール内呼称
    var emailAddress = getRange(mySheet,i,mailAddressCol).getValue();　
    if (documentId.length == 0 || hisName.length == 0 ||  emailAddress.length == 0){
      continue;
    }

    var body = getMailTemplateBody(mailTemplate);
    // 新しい本文を生成 (ここで置換を全部やる)
    var newBody=body
      .replaceText("{お名前}",hisName)
      .replaceText("{ドキュメント}",fileUrl);
    
    // リンクを編集
    var mailBody = replaceLink(newBody, fileUrl).getText();
    
    var to = emailAddress;
    var subject = title
      .replace(/{お名前}/,hisName)
      .replace(/{ドキュメント}/,fileUrl);    
    
    // 生成した本文をメールで送信  
    GmailApp.sendEmail(
      to,
      subject,
      mailBody
    ); //MailAppではfromが設定できないとのこと
    Logger.log("メールを送信しました：" + to); //ドキュメントの内容をログに表示
//     Logger.log(newBody.getText());
  }
   
}

function replaceLink(body, url){
    var urlLink = null;
    
    while (urlLink = body.findText(url, urlLink)){
      var originUrl = urlLink.getElement().asText().getLinkUrl();
//      if(originUrl == null && urlLink.isPartial()){        
//        continue;
//      }
      Logger.log("リンクを設定します:" & urlLink.getElement().asText());
      
      urlLink.getElement().asText().setLinkUrl(url);
    }
    return body;
}

function generateShortUrls(){
  Logger.log("スプレッドシート内のURLを短縮します。");

  /* スプレッドシートのシートを取得と準備 */
  var mySheet=SpreadsheetApp.getActiveSheet(); //シートを取得
  var rowSheet = mySheet.getDataRange().getLastRow(); //シートの使用範囲のうち最終行を取得
  for(var i = 2 ; i <= rowSheet; i++ ){
    var personName = getRange(mySheet,i,NAME_COL).getValue().replace(" ", "");
    if (personName.length == 0){
      Logger.log("終了します。");
      break;
    }
  
    var newUrlCol = getNewUrlCol();
    var originUrlCol = getOriginUrlCol();
    
    var currentValue = getRange(mySheet,i,newUrlCol).getValue();
    if( currentValue != ""){
      continue;
    }
    
    var originUrl = getRange(mySheet,i,originUrlCol).getValue();　// 元のURL
    var shortUrl =  shorten(originUrl);
    getRange(mySheet,i,newUrlCol).setValue(shortUrl);
  }
}

function shorten(originUrl){
   //var url = UrlShortener.Url.insert({longUrl: originUrl});
   //return url.id; 
  
   var token = PropertiesService.getDocumentProperties().getProperty("bitly_token");
   var url = "https://api-ssl.bitly.com/v3/shorten?access_token=" + token + "&longUrl=" + originUrl;
   var responseApi = UrlFetchApp.fetch(url);
   var responseJson = JSON.parse(responseApi.getContentText());
   return responseJson["data"]["url"];
}

function generateFiles(folder){
  
  /* スプレッドシートのシートを取得と準備 */
  var mySheet=SpreadsheetApp.getActiveSheet(); //シートを取得
  var rowSheet=mySheet.getDataRange().getLastRow(); //シートの使用範囲のうち最終行を取得
  var docIdCol = getNewIdCol();

  /* テンプレートは独立した文書で、ひとつだけ使う */
  var props = PropertiesService.getDocumentProperties();
  var strDocUrl= props.getProperty("templateDocId"); //ドキュメントのURL
  var templateFile = DriveApp.getFileById(strDocUrl); 
    
  /* シートの全ての行について社名、姓名を差し込みログに表示*/
  for(var i=2;i<=rowSheet;i++){
    var number = "000" + getRange(mySheet,i,1).getValue();
    number = number.substring(number.length - 3);
    var personName = getRange(mySheet,i,NAME_COL).getValue().replace(" ", "");
    if (personName.length == 0){
      break;
    }
    var fileName = number + "_" + personName + "さん";

    var newFile = templateFile.makeCopy(fileName, folder);
//    var docIdCol = 
    getRange(mySheet,i,docIdCol).setValue(newFile.getId()); // 新しいドキュメントIDを控えておく
    var newDocument=DocumentApp.openById(newFile.getId()); //ドキュメントをIDで取得
    var body = newDocument.getBody();
 
    var newUrlCol = getNewUrlCol(); 
    var shortUrl = getRange(mySheet,i,newUrlCol).getValue();　// 短縮URL
    var partnerName =  getRange(mySheet,i,NAME_COL).getValue();　// 紹介者名
//    var strMessage =mySheet.getRange(i,3).getValue();　//メッセージ 

    // 新しい本文を生成 (ここで置換を全部やる)
    var newBody=body
      .replaceText("{紹介者}",partnerName)
      .replaceText("{短縮URL}",shortUrl);

    // リンクを編集
    replaceLink(newBody, shortUrl);
    
    Logger.log("書き込みました：" + newBody); //ドキュメントの内容をログに表示
 
  }
}

// 水平線の手前まで、ドキュメントの内容を取得します。
function getTemplateSection(document){
    var searchTypeParagraph = DocumentApp.ElementType.PARAGRAPH;
    var searchTypeHR = DocumentApp.ElementType.HORIZONTAL_RULE;
    var body = document.getBody();
    var firstHR = body.findElement(searchTypeHR);
    var templateBody = "";
    var templateParagraph = null;
    var theHr = firstHR.getElement();
    Logger.log(theHr.getParent());
    
    // 水平線(HR)の前までがテンプレート。これを取得して値を差し込み、新しい本文を作る
    while (templateParagraph = body.findElement(searchTypeParagraph, templateParagraph)){
      var theParagraph = templateParagraph.getElement().asParagraph();
      Logger.log("既存のテンプレート：" + theParagraph.getText());
//      Logger.log("水平線前の段落" + theHr.getParent().getText());

      if(body.getChildIndex( theHr.getParent()) <body.getChildIndex(theParagraph)){
        Logger.log("水平線が見つかりました。" );
        break;
      }
      templateBody += theParagraph.getText() + "\n"; //最初の段落の内容を取得
    }
    Logger.log("テンプレート全体： " + templateBody); 
    return templateBody;
}

function createNewFolder(){
  // 出力先のフォルダを生成
  Logger.log("出力先のフォルダを生成");
  
  var sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();
  Logger.log("スプレッドシートID：" + sheetId);
  var file = DriveApp.getFileById(sheetId);
  var thisFolder  = file.getParents().next();
  var today = new Date(); 
  var dateString = "";
  dateString += today.getFullYear() + "-";
  dateString += (today.getMonth() + 1) + "-";
  dateString += today.getDate();
  
  while(thisFolder.getFoldersByName(dateString).hasNext()){
    var child = thisFolder.getFoldersByName(dateString).next();
    child.setTrashed(true);
    Logger.log("削除しました：" + child.getId());  
  }
  
  var newFolder = thisFolder.createFolder(dateString);
  newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.EDIT);

  Logger.log("生成しました：" + dateString);
  
  return newFolder;
}

// 過去の出力結果を削除します。
function removeOldOutput(document){

  

}
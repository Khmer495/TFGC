var sourcefileName = 't55-73_8';
var title = 'Special G-2';
var mailTo = '';
var mailSubject = 'TFGC：THINK FUTURE Grammar Challenge ' + title + "\n";
var mailBody = 'QRコード：THINK FUTURE Grammar Challenge ' + title + "\n";

var TFGCRootFolderID = DriveApp.getFoldersByName('TFGC').next().getId();
var TFGCSourceFolderID = DriveApp.getFolderById(TFGCRootFolderID).getFoldersByName('source').next().getId();
var sourceSheet = SpreadsheetApp.openById(DriveApp.getFolderById(TFGCSourceFolderID).getFilesByName(sourcefileName).next().getId()).getSheets()[0];
var TFGCFormFolderID = DriveApp.getFolderById(TFGCRootFolderID).getFoldersByName('form').next().getId();
var TFGCQRFolderID = DriveApp.getFolderById(TFGCRootFolderID).getFoldersByName('QR').next().getId();
var TFGCDocumentFolderID = DriveApp.getFolderById(TFGCRootFolderID).getFoldersByName('document').next().getId();
var fileName = 'TFGC_' + sourcefileName;

function makeTFGC(){
  var urlimg = makeForm();
  var pdf = makeDoc();
  sendMail(urlimg[0], urlimg[1], pdf);
}

function makeForm(){
  
  Logger.log('start making form');
  
  var form = FormApp.create(fileName);
  var formID = form.getId();
  form.setTitle('THINK FUTURE Grammar Challenge ' + title);
  form.setDescription('入力必須項目、一部該当者のみの入力項目、問題の解答ページの順となっています。');
  form.addSectionHeaderItem().setTitle('入力必須項目');
  form.addTextItem()
    .setTitle('苗字をカタカナで入力してください。')
    .setHelpText('（例）ヤマダ')
    .setRequired(true)
  ;
  form.addTextItem()
    .setTitle('名前をカタカナで入力してください。')
    .setHelpText('（例）タロウ')
    .setRequired(true)
  ;
  Logger.log('complete first page');
  
  form.addPageBreakItem()
    .setTitle('以下の項目について、以前入力したことがある方は飛ばしてください。')
    .setHelpText('変更がある場合は該当箇所のみ入力してください。')
  ;
  form.addTextItem()
    .setTitle('苗字を漢字で入力してください。')
    .setHelpText('（例）山田')
  ;
  form.addTextItem()
    .setTitle('名前を漢字で入力してください。')
    .setHelpText('（例）太郎')
  ;
  form.addTextItem()
    .setTitle('在籍する高校を教えてください。')
    .setHelpText('～高校という形式で入力してください。（例：三国丘高校）')
    .setValidation(FormApp.createTextValidation()
                   .requireTextContainsPattern('高校')
                   .setHelpText('～高校という形式で入力してください。')
                   .build()
                   )
  ;
  form.addMultipleChoiceItem()
    .setTitle('学年を教えてください。')
    .setChoiceValues(range(1, 4))
  ;
  form.addTextItem()
    .setTitle('志望大学はどこですか。')
    .setHelpText('～大学という形式で答えてください。（例：大阪大学）')
    .setValidation(FormApp.createTextValidation()
                   .requireTextContainsPattern('大学')
                   .setHelpText('～大学という形式で入力してください。')
                   .build()
                   )
  ;
  form.addTextItem()
    .setTitle('志望学部・学域はどこですか。')
    .setHelpText('～学部または学域という形式で答えてください。（例：工学部、工学域）')
  ;
  Logger.log('complete second page');
  
  form.addPageBreakItem()
    .setHelpText('次のページから解答ページに移ります。')
  ;
  Logger.log('complete third page');
  
  form.addPageBreakItem();
  
  var header = 1;
  var direction = '';
  var question = '';
  var choice_array = [];
  var choice = '';
  var choice_len = 0;
  
  for(var i = 1 + header; i <= sourceSheet.getLastRow(); i++){
    direction = sourceSheet.getRange(i,5).getValue().replace('\\\\', '\n\t');
    question = sourceSheet.getRange(i,6).getValue().replace('\\\\', '\n\t');
    choice_array = sourceSheet.getRange(i,7).getValue().split('\\\\')
    choice = ''
    for(var j = 0; j < choice_array.length; j++){
      if(j < choice_array.length - 1){
        choice += '\t\t' + (j+1) + '. ' + choice_array[j] + '\n';
      }else{
        choice += '\t\t' + (j+1) + '. ' + choice_array[j];
      }
    }
    form.addSectionHeaderItem().setHelpText('(' + (i-1) + ')\t' + direction + '\n\t' + question + '\n' + choice);
    choice_len = choice_array.length;
    form.addMultipleChoiceItem()
      .setTitle('(' + (i - header) + ')')
      .setChoiceValues(range(1, choice_len + 1))
      .setRequired(true)
    ;
    Logger.log('complete forth page, (' + (i - header) + ')');
  }
  
  var url = form.shortenFormUrl(form.getPublishedUrl());
 
  DriveApp.getFolderById(TFGCFormFolderID).addFile(DriveApp.getFileById(formID));
  DriveApp.getRootFolder().removeFile(DriveApp.getFileById(formID));
  Logger.log('complete making form');
  
  var img = makeQR(url);
  
  return [url, img];
}

function makeQR(url){
  Logger.log('start making QR');
  Logger.log('url:' + url);
  
  var img = UrlFetchApp.fetch('https://chart.googleapis.com/chart?chs=200x200&cht=qr&chl=' + url).getBlob().setName(fileName + '.png');
  DriveApp.getFolderById(TFGCQRFolderID).createFile(img);
  
  Logger.log('complete making QR');
  
  return img;
}

function sendMail(url, img, pdf){
  MailApp.sendEmail({
    to: mailTo,
    subject: mailSubject,
    body: mailBody + "url: " + url,
    attachments: [img, pdf]
  });
}

function makeDoc() {
  
  var fontpt = 11;
  var margin = fontpt * 2;
  
  var documentFile = DocumentApp.create(fileName);
  var documentID = documentFile.getId();
  var document = documentFile.getBody()
    .setMarginTop(margin)
    .setMarginLeft(margin)
    .setMarginRight(margin)
    .setMarginBottom(margin)
  ;
  var docTitle = 
      'THINK FUTURE Grammar Challenge ' + title + '\n'
      //+ '\n'
      //+ '問題は全部で' + (sourceSheet.getLastRow()-1) + '問あります。\n'
      //+ '右のQRコードを読み取り、Googleフォームから解答してください。\n'
      //+ 'すべて解いてから提出してください。\n'
      ;
  
  var p_title = document.appendParagraph('')
    .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
    .setLineSpacing(1) 
  ;
  p_title.appendText(docTitle);
  Logger.log('set title');
  
  var direction = '';
  var question = '';
  var choice_array = [];
  var choice = '';
  var text = '';
  var under_start = 0;
  var under_end = 0;
  var under_bar_re = '_(.*?)_';
  var p_question = '';
  var p_choice = ''
  for(var i = 2; i <= sourceSheet.getLastRow(); i++){
    Logger.log('set question ' + (i-1));
    direction = sourceSheet.getRange(i,5).getValue().replace('\\\\', '\n');
    question = sourceSheet.getRange(i,6).getValue().replace('\\\\', '\n');
    choice_array = sourceSheet.getRange(i,7).getValue().split('\\\\');
    choice = ''
    
    p_question = document.appendParagraph('')
      .setIndentFirstLine(0)
      .setIndentStart(fontpt * 2)
      .setLineSpacing(1)
    ;
    p_question.appendText('(' + (i-1) + ')\t' + direction + '\n');
    p_question.appendText(question);
    
    for(var j = 0; j < choice_array.length; j++){
      if(j < choice_array.length - 1){
        choice += (j+1) + '. ' + choice_array[j] + '      ';
      }else{
        choice += (j+1) + '. ' + choice_array[j] + '\n';
      }
    }
    p_choice = document.appendParagraph('')
      .setIndentFirstLine(fontpt * 4)
      .setIndentStart(fontpt * 4)
      .setLineSpacing(1)
    ;
    p_choice.appendText(choice);
  }
  while(document.editAsText().findText(under_bar_re) != null){
    p = document.editAsText().findText(under_bar_re);
    under_start = p.getStartOffset();
    under_end = p.getEndOffsetInclusive();
    p.getElement().asText().setUnderline(under_start, under_end, true);
    p.getElement().asText().deleteText(under_end, under_end);
    p.getElement().asText().deleteText(under_start, under_start);
    if(p.getElement().asText().getText()[under_start-1] = '(\d)'){
      p.getElement().asText().setTextAlignment(under_start-1,under_start-1, DocumentApp.TextAlignment.SUBSCRIPT);
    }
  }
  
  documentFile.saveAndClose();
        
  DriveApp.getFolderById(TFGCDocumentFolderID).addFile(DriveApp.getFileById(documentID));
  DriveApp.getRootFolder().removeFile(DriveApp.getFileById(documentID));
  Logger.log('complete making document');
  
  var pdf = documentFile.getAs('application/pdf');
  DriveApp.getFolderById(TFGCDocumentFolderID).createFile(pdf).setName(fileName);
  Logger.log('complete making pdf');
  
  return pdf;
}

function openSpreadsheetByName(name) {

  // get a collection of all files
  var files = DriveApp.getRootFolder().getFilesByName(name + '.csv').next;
  
  // get file
  var file = files.next();
  
  // open spreadsheet
  return SpreadsheetApp.openById(file.getId());
}

function range(from, to) {
  var array = [];
  for (var i = from; i < to; i++) {
    array.push(i)
  }
  return array;
}

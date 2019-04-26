// ダイアログから入力した文字列を、用語集に追加するスクリプト
function addGlossary(){
  var input = Browser.inputBox('用語集に追加 [単語,訳語（省略可）]');
  
  // 「✕」を押した時に返ってくる「cancel」を除外
  if(input == "cancel"){
    return;
  }

  var words = input.split(",");
  var word_en = words[0];
  if(words.length > 1){
    var word_jp = words[1];
  }else{
    var word_jp = "";
  }
  
  // スプレッドシートの読み込み  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var glossary_sheet = spreadsheet.getSheetByName(sheetName_glossary);
  
  var en_range = column_glossary_en + first_row_texts + ":" + column_glossary_en;
  var jp_range = column_glossary_jp + first_row_texts + ":" + column_glossary_jp;
 
  var glossary_en_texts = glossary_sheet.getRange(en_range).getValues();
  var glossary_jp_texts = glossary_sheet.getRange(jp_range).getValues();
  
  var glossary_en = glossary_sheet.getRange(en_range);
  var glossary_jp = glossary_sheet.getRange(jp_range);
  
  var last_row = glossary_en_texts.filter(String).length;

  glossary_en.getCell(last_row+1,1).setValue(word_en);
  glossary_jp.getCell(last_row+1,1).setValue(word_jp);

}
// 用語集の正規表現リストを作る
function getGlossaryReg(){
  
  // スプレッドシートの読み込み  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var glossary_sheet = spreadsheet.getSheetByName(sheetName_glossary);
  
  var glossary_last_row = glossary_sheet.getLastRow();
  var glossary_en_range = column_glossary_en + first_row_glossary + ":" + column_glossary_en + glossary_last_row;
  var glossary_jp_range = column_glossary_jp + first_row_glossary + ":" + column_glossary_jp + glossary_last_row;

  var glossary_en = glossary_sheet.getRange(glossary_en_range).getValues();
  var glossary_jp = glossary_sheet.getRange(glossary_jp_range).getValues();
  
  var glossary_reg = [];
  for(var row_glossary = 0; row_glossary < glossary_en.length; row_glossary++){

    // 原語
    var word_en = glossary_en[row_glossary][0].toString();
    if(word_en == ""){
      continue;
    }

    // 訳語
    var word_jp = glossary_jp[row_glossary][0].toString();
    if(word_jp == ""){
      continue;
    }
    
    // 正規表現
    var re_en = new RegExp(word_en, "ig"); // i:大文字小文字を区別しない、g:全てをマッチング
    var re_jp = new RegExp(word_jp, "ig"); // i:大文字小文字を区別しない、g:全てをマッチング
    
    // 置換用のリスト
    glossary_reg.push([re_en,re_jp,word_en,word_jp]);
  }
 
  // 用語集の原語の文字数が多い順にソートする（部分一致を避けるため）
  glossary_reg.sort(function(a,b){
    return b[1].length-a[1].length;
  });
  
  return glossary_reg;
}
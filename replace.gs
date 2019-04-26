// テキスト内の固有名詞原語を、用語集シートの内容で置換するスクリプト
function replaceByGlossary(sheetName_replace, column_replace_input, column_replace_output){

  // スプレッドシートの読み込み  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = spreadsheet.getSheetByName(sheetName_replace);
  var last_row = sheet.getLastRow();
  var input_range = column_replace_input + first_row_texts + ":" + column_replace_input + last_row;
  var output_range = column_replace_output + first_row_texts + ":" + column_replace_output + last_row;
  
  var input_texts = sheet.getRange(input_range).getValues();
  var output_texts = sheet.getRange(output_range).getValues();
  var output = sheet.getRange(output_range);

  var glossary_reg = getGlossaryReg();
  
  // 用語の置換
  for(var row_input = 0; row_input < input_texts.length; row_input++){
    
    var input_text = input_texts[row_input][0].toString();
    if(input_text == ""){
      continue;
    }
    
    // 置換
    for(var row_glossary = 0; row_glossary < glossary_reg.length; row_glossary++){
      input_text = input_text.replace(glossary_reg[row_glossary][0], glossary_reg[row_glossary][3]);
    }
    
    // 結果のテキストを保存
    output_texts[row_input][0] = input_text;
  }
  
  // 結果を出力
  output.setValues(output_texts);
  
  // 色つけ
  highlightReplaced();
}

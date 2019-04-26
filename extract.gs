// テキスト内の固有名詞原語を、抽出するスクリプト
function extractByGlossary(sheetName_extract, column_extract_input, column_extract_output){
  // スプレッドシートの読み込み  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = spreadsheet.getSheetByName(sheetName_extract); 
  var last_row = sheet.getLastRow();
  var input_range = column_extract_input + first_row_texts + ":" + column_extract_input + last_row;
  var output_range = column_extract_output + first_row_texts + ":" + column_extract_output + last_row;
  
  var input_texts = sheet.getRange(input_range).getValues();
  var output_texts = sheet.getRange(output_range).getValues();
  var output = sheet.getRange(output_range);

  var glossary_reg = getGlossaryReg();
  
  // 用語の抽出
  for(var row_input = 0; row_input < input_texts.length; row_input++){
    var input_text = input_texts[row_input][0].toString();
    
    if(input_text == ""){
      continue;
    }

    var extracted_words = [];
    
    for(var row_glossary = 0; row_glossary < glossary_reg.length; row_glossary++){
      if(input_text.match(glossary_reg[row_glossary][0]) !== null)
      {
        extracted_words.push(glossary_reg[row_glossary][2] + "," + glossary_reg[row_glossary][3]);
        
        // 一致したものは削除(部分一致を避けるため)
        input_text = input_text.replace(glossary_reg[row_glossary][0], "");
      }
    }
    
    // 抽出した用語を整形
    var result = "";
    for(var num = 1; num <= extracted_words.length; num++){
      result += num + "," + extracted_words[num-1];
      if(num < extracted_words.length){
        result += String.fromCharCode(10); // 改行
      }
    }
    output_texts[row_input][0] = result;

  }

  // 結果を出力
  output.setValues(output_texts);
  
  // 色つけ
  highlightExtract();
  
}

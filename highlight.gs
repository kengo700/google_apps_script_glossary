// テキスト内の固有名詞原語を、ハイライトするスクリプト
function highlightByGlossary(sheetName_highlight, column_highlight_input, column_highlight_output, color_en, color_jp){
  
  // スプレッドシートの読み込み  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  
  var sheet = spreadsheet.getSheetByName(sheetName_highlight);
  var last_row = sheet.getLastRow();
  var input_range = column_highlight_input + first_row_texts + ":" + column_highlight_input + last_row;
  var output_range = column_highlight_output + first_row_texts + ":" + column_highlight_output + last_row;
  
  var input_texts = sheet.getRange(input_range).getValues();
  var output = sheet.getRange(output_range);
  
  var glossary_reg = getGlossaryReg();
    
  // ハイライトのスタイル設定
  var style = SpreadsheetApp.newTextStyle();
  style.setForegroundColor(color_en);
  var buildStyle_en = style.build();
  style.setForegroundColor(color_jp);
  var buildStyle_jp = style.build();
  
  var rich = SpreadsheetApp.newRichTextValue();
  
  // 用語のハイライト
  for(var row_input = 0; row_input < input_texts.length; row_input++){
    var input_text = input_texts[row_input][0].toString();
    
    if(input_text == ""){
      continue;
    }
    
    rich.setText(input_text);
    
    for(var row_glossary = 0; row_glossary < glossary_reg.length; row_glossary++){

      // 英語に色つけ
      while ((result = glossary_reg[row_glossary][0].exec(input_text)) !== null)
      {  
        var start = result.index;
        var end = start + result[0].length - 1;
        Logger.log(start + " " + end);

        rich.setTextStyle(start,end+1,buildStyle_en);
      }
      
      // 日本語に色つけ
      while ((result = glossary_reg[row_glossary][1].exec(input_text)) !== null)
      {  
        var start = result.index;
        var end = start + result[0].length - 1;
        Logger.log(start + " " + end);

        rich.setTextStyle(start,end+1,buildStyle_jp);
      }

    }
  
    var format = rich.build()
    
    // 結果を保存
    output.getCell(row_input+1,1).setRichTextValue(format);
  }
  
}
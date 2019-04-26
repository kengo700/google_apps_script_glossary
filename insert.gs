// 選択中のセルの行の、抽出した用語のX番目を、訳文の最後に追加するスクリプト
function insertWord(sheetName_insert, column_insert_input, column_insert_output, num){
  // スプレッドシートの読み込み  
  var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = spreadsheet.getSheetByName(sheetName_insert);
  
  var active_cell = sheet.getActiveCell();
  var row = active_cell.getRow();

  var input_cell = column_insert_input + row;
  var output_cell =  column_insert_output + row;
  var extracted_words = sheet.getRange(input_cell).getValue();
  
  if(extracted_words==""){
    return;
  }
  
  // １行ずつに分割
  var extracted_words_lines = extracted_words.split(/\r\n|\r|\n/);
  
  // 挿入する用語を取得
  if(num-1 < extracted_words_lines.length){
    var insert_word = extracted_words_lines[num-1].split(",")[1];
  }else{
    var insert_word = "";
  }

  // 挿入
  var text = sheet.getRange(output_cell).getValue();
  sheet.getRange(output_cell).setValue(text + insert_word)
  
}
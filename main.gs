// 翻訳データシートの設定
var sheetName_texts = "翻訳データ";
var first_row_texts = 2; // １行目は見出し
var column_texts_en = "A";
var column_texts_jp = "B";
var column_texts_extract = "D";
var column_texts_replace = "E";

// 用語集シートの設定
var sheetName_glossary = "用語集";
var first_row_glossary = 2; // １行目は見出し
var column_glossary_en = "B";
var column_glossary_jp = "C";

// 更新
function update(){
  highlightOriginal();
  highlightTranslated();
  extractOriginal();
  replaceTranslated();
}

// 訳文内の固有名詞を置換して出力
function replaceTranslated(){
  replaceByGlossary(sheetName_texts, column_texts_jp, column_texts_replace);
}  

// 原文内の固有名詞原語を抽出する
function extractOriginal(){
 extractByGlossary(sheetName_texts, column_texts_en, column_texts_extract);
}
  
// 原文中の用語をハイライト
function highlightOriginal(){
    highlightByGlossary(sheetName_texts, column_texts_en, column_texts_en,"#FA5858", "#5882FA");
}

// 訳文中の用語をハイライト
function highlightTranslated(){
    highlightByGlossary(sheetName_texts, column_texts_jp, column_texts_jp,"#FA5858", "#5882FA");
}

// 置換後の日本語をハイライト
function highlightReplaced(){
    highlightByGlossary(sheetName_texts, column_texts_replace, column_texts_replace,"#FA5858", "#5882FA");
}

// 抽出した用語をハイライト
function highlightExtract(){
    highlightByGlossary(sheetName_texts, column_texts_extract, column_texts_extract,"#FA5858", "#5882FA");
}

// 抽出したテキストのX番目を挿入（マクロに登録してショートカットから実行できるように、1~9の関数を用意）
function insertWord1(){
  insertWord(sheetName_texts, column_texts_extract, column_texts_jp, 1);
}
function insertWord2(){
  insertWord(sheetName_texts, column_texts_extract, column_texts_jp, 2);
}
function insertWord3(){
  insertWord(sheetName_texts, column_texts_extract, column_texts_jp, 3);
}
function insertWord4(){
  insertWord(sheetName_texts, column_texts_extract, column_texts_jp, 4);
}
function insertWord5(){
  insertWord(sheetName_texts, column_texts_extract, column_texts_jp, 5);
}
function insertWord6(){
  insertWord(sheetName_texts, column_texts_extract, column_texts_jp, 6);
}
function insertWord7(){
  insertWord(sheetName_texts, column_texts_extract, column_texts_jp, 7);
}
function insertWord8(){
  insertWord(sheetName_texts, column_texts_extract, column_texts_jp, 8);
}
function insertWord9(){
  insertWord(sheetName_texts, column_texts_extract, column_texts_jp, 9);
}




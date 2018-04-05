function sendArticleCountEmails(){
  //「Send-Email」シートの A3:B100 のデータを取得し、edata 配列（edata[行][列]）に格納
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.setActiveSheet(ss.getSheetByName("Send-Emails"));
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange("A3:B100");
  var edata = dataRange.getValues();

  //「TOTAL Summary」シートの A1:Z19 のデータを取得し、data 配列（edata[行][列]）に格納
  // 実際には、２つのサイトなら、A1:C19、３つなら A1:D19 でOK。ここではZぐらいまで多数入れたと仮定してます。 
  var dd = SpreadsheetApp.getActiveSpreadsheet();
  dd.setActiveSheet(dd.getSheetByName("TOTAL Summary"));
  var sheet = SpreadsheetApp.getActiveSheet();
  var dataRange = sheet.getRange("A1:Z19");
  var data = dataRange.getValues();

  // 1-9 の数字を 01-09 に直します。 
  var toDoubleDigits = function(num) {
    num += "";
    if (num.length === 1) {
      num = "0" + num;
    }
    return num;     
  };

  //「TOTAL Summary」シートの一行目から、データ取得の範囲を抽出します。それを YYYY-MM-DD 形式に変換します。
  var from_date_n = new Date(data[0][1]);
  var to_date_n = new Date(data[0][3]);
  var from_date = from_date_n.getFullYear()+'-'+toDoubleDigits(from_date_n.getMonth()+1)+'-'+toDoubleDigits(from_date_n.getDate());
  var to_date = to_date_n.getFullYear()+'-'+toDoubleDigits(to_date_n.getMonth()+1)+'-'+toDoubleDigits(to_date_n.getDate());

  // スプレットシートを PDF取得（Google-Analytics-Reports-20180401_20180404.pdf のように開始日と終了日がファイル名に入るようにしておきます）
  //var pdf = dd.getAs('application/pdf').setName("Google-Analytics-Reports-"+from_date+"_"+to_date+".pdf"); 
  var ssid = dd.getId();
  var sheetid = sheet.getSheetId();
  var filename = "CSEAS-Google-Analytics-Reports-"+from_date+"_"+to_date;
  var pdf = createPDF(ssid, sheetid, filename);
  
  // E-mailデータの抽出
  for(i in edata){
    if(!edata[i][0]) break;
    var rowData = edata[i];
    var emailAddress = rowData[1];
    var recipient = rowData[0];
    var message ='Dear '+ recipient +',\n\n';
    var subject ='Google Analytics Logs from '+from_date+' to '+to_date;

    // HTML本文を htmlbody に格納していきます。
    // 基本 table タグのみで構成してみました。border はメール本文だと鬱陶しかったので指定してません。
    var htmlbody = "<table>\n";
    data[0][1] = from_date;
    data[0][3] = to_date;
    for(j in data){
      htmlbody = htmlbody + "<tr>\n";
      for(k in data[j]){
        if(data[j][1] == "") continue;
        htmlbody = htmlbody + "  <td>" + data[j][k] + "</td>\n";
      }
    }
    // メール送信（htmlbody に入れたHTMLと、pdf に格納していた添付ファイルを指定してます。HTMLメールを受け付けない場合には、「Please allow a HTML message.」を表示しておきます。）
    MailApp.sendEmail(emailAddress, subject, 'Please allow a HTML message.', {htmlBody:htmlbody,attachments: [pdf]});
  }
}

// https://www.virment.com/create-pdf-google-apps-script/#PDF を参考に一部変更しました。
// PDF作成関数　引数は（ssid:PDF化するスプレッドシートID, sheetid:PDF化するシートID）
function createPDF(ssid, sheetid, filename){
  
  // スプレッドシートをPDFにエクスポートするためのURL。このURLに色々なオプションを付けてPDFを作成
  var url = "https://docs.google.com/spreadsheets/d/SSID/export?".replace("SSID", ssid);
 
  // PDF作成のオプションを指定
  var opts = {
    exportFormat: "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
    format:       "pdf",    // ファイル形式の指定 pdf / csv / xls / xlsx
    size:         "A4",     // 用紙サイズの指定 legal / letter / A4
    portrait:     "false",   // true → 縦向き、false → 横向き
    fitw:         "true",   // 幅を用紙に合わせるか
    sheetnames:   "false",  // シート名をPDF上部に表示するか
    printtitle:   "false",  // スプレッドシート名をPDF上部に表示するか
    pagenumbers:  "false",  // ページ番号の有無
    gridlines:    "false",  // グリッドラインの表示有無
    fzr:          "false",  // 固定行の表示有無
    gid:          sheetid   // シートIDを指定 sheetidは引数で取得
  };
  
  var url_ext = [];
  
  // 上記のoptsのオプション名と値を「=」で繋げて配列url_extに格納
  for( optName in opts ){
    url_ext.push( optName + "=" + opts[optName] );
  }
 
  // url_extの各要素を「&」で繋げる
  var options = url_ext.join("&");
 
  // API使用のためのOAuth認証
  var token = ScriptApp.getOAuthToken();
 
  // PDF作成
  var response = UrlFetchApp.fetch(url + options, {
      headers: {
        'Authorization': 'Bearer ' +  token
      }
  });

  // ファイル名の指定
  var pdf = response.getBlob().setName(filename + '.pdf');
  
  return pdf;
 
}
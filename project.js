//----------------------------API Connection--------------------------------//
var token = "6015555381:AAEYKVe0QF8-q_Qt26WW6C491khB_VkwMTw";
var SheetID = "139PfQrM61tv214yUCbRR_lI6hfG_GtSYPJ2BR7viKsg";
var url = "https://api.telegram.org/bot" + token;
var webAppUrl = "https://script.google.com/macros/s/AKfycbxQe8Dw4RhPFmYgIZD7nnVAruoTi-rMnFb59Ewtj8ZA951JGKo4w2WcvEcZT2C98V9h/exec";

function setWebhook() {
    var response = UrlFetchApp.fetch(url + "/setWebhook?url=" + webAppUrl);
    Logger.log(response.getContentText());
}

//-----------------------------Command List--------------------------------//
function doPost(e) {
  var stringJson = e.postData.getDataAsString();
  var updates = JSON.parse(stringJson);
  
  var firstName = updates.message.from.first_name;
  var lastName = updates.message.from.last_name;
  var id = updates.message.chat.id;
  var textBot = updates.message.text;
  var captionBot = updates.message.caption;

  if (textBot || captionBot) {
    var groupName = updates.message.chat.title;
    var chat_bot = textBot || captionBot;
    var command_cek = chat_bot.substring(0, 1);
    var command = chat_bot.split(" ")[0];
    var args = chat_bot.split(" ")[1];

    if (command_cek == "/") {
      switch (command) {
        case "/start":
          let starting = "<b>Halo !! Selamat Datang di Filter bot IOC-W ASSURANCE</b> \n" +
            "Untuk memulai cek perintah di /help";
          sendText(id, starting);
          break;
        case "/help":
          let help = "<b>Halo !! Berikut Panduan Command untuk melakukan pencarian data :</b>\n\n" +
            "/statusHELP Cek perintah yang tersedia untuk STATUS \n" +
            "/stoHELP Cek perintah yang tersedia untuk STO \n" +
            "/detailHELP Cek perintah yang tersedia untuk Detail \n" +
            "/report Cek reporting \n" +
            "/Input Melakukan input laporan \n" +
            "/Close Melakukan input lapuran untuk rekap";
          sendText(id, help);
          break;
        case "/report":
          sendText(id, sendReporting());
          break;
        
        //-----------------------------Command List INPUT--------------------------------//
        case "/Input":
          var dataToInput = chat_bot.split("\n");
          if (dataToInput.length == 8) {
            var reportDate = getCurrentDateTime();
            var inputer = firstName;
            if (lastName) {
              inputer += " " + lastName;
            }
            var namaPelanggan = dataToInput[0].split(": ")[1];
            var contact = dataToInput[1].split(": ")[1];
            var noUser = dataToInput[2].split(": ")[1];
            var datel = dataToInput[3].split(": ")[1];
            var customerType = dataToInput[4].split(": ")[1];
            var alamatcust = dataToInput[5].split(": ")[1];
            var laporanGangguan = dataToInput[6].split(": ")[1];
            var sto = dataToInput[7].split(": ")[1];

            writeToSpreadsheet(id, reportDate, inputer, namaPelanggan, contact, noUser, datel, customerType, alamatcust, laporanGangguan, sto);

          } else {
            let notif = '<b>***Template Format Input Laporan***</b> \n' +
              '/Input Nama Pelanggan: \n' +
              'Contact: \n' +
              'No User: \n' +
              'DATEK: \n' +
              'Customer Type: \n' +
              'Alamat Customer: \n' +
              'Laporan Gangguan: \n' +
              'STO: \n\n' +
              '/help';
            sendText(id, notif);
          }
          break;
        case "/Close":
          var dataToInput = chat_bot.split("\n");
          if (dataToInput.length == 32) {
            var reportDate = getCurrentDateTime();
            var inputer = firstName;
            if (lastName) {
              inputer += " " + lastName;
            }
            var close = dataToInput[0].split(": ")[1];
            var tiket = dataToInput[1].split(": ")[1];
            var jenisGangguan = dataToInput[2].split(": ")[1];
            var sektor = dataToInput[4].split(": ")[1];
            var sto = dataToInput[5].split(": ")[1];
            var user = dataToInput[6].split(": ")[1];
            var odp = dataToInput[7].split(": ")[1];
            var rfo = dataToInput[8].split(": ")[1];
            var penyebabGangguan = dataToInput[9].split(": ")[1];
            var dropcore = dataToInput[12].split(": ")[1];
            var soc = dataToInput[13].split(": ")[1];
            var roset = dataToInput[14].split(": ")[1];
            var adaptor = dataToInput[15].split(": ")[1];
            var sleeve = dataToInput[16].split(": ")[1];
            var otp = dataToInput[17].split(": ")[1];
            var prekso = dataToInput[18].split(": ")[1];
            var heatshink = dataToInput[19].split(": ")[1];
            var dll = dataToInput[20].split(": ")[1];
            var ontLama = dataToInput[23].split(": ")[1];
            var ontBaru = dataToInput[24].split(": ")[1];
            var stbLama = dataToInput[25].split(": ")[1];
            var stbBaru = dataToInput[26].split(": ")[1];
            var teknisi1 = dataToInput[29].split(": ")[1];
            var teknisi2 = dataToInput[30].split(": ")[1];
            var teknisi3 = dataToInput[31].split(": ")[1];

            writeClose(groupName, id, reportDate, inputer, close, tiket, jenisGangguan, sektor, sto, user, odp, rfo, penyebabGangguan, dropcore, soc, roset, adaptor, sleeve, otp, prekso, heatshink, dll, ontLama, ontBaru, stbLama, stbBaru, teknisi1, teknisi2, teknisi3);

          } else {
            let notif = '<b>***Template Format Input Laporan***</b> \n' +
              '/Close : \n' +
              'TIKET: \n' +
              'JENIS GANGGUAN: \n\n' +
              'SEKTOR: \n' +
              'STO: \n' +
              'USER: \n' +
              'ODP: \n' +
              'RFO: \n' +
              'PENYEBAB GANGGUAN: \n\n' +
              '<b>==MATERIAL==</b> \n'+
              'DROPCORE: \n' +
              'SOC: \n' +
              'ROSET: \n' +
              'ADAPTOR: \n' +
              'SLEEVE: \n' +
              'OTP: \n' +
              'PREKSO: \n' +
              'HEATSHINK: \n' +
              'DLL: \n\n'+
              '<b>==JIKA PERGANTIAN NTE==</b> \n'+
              'ONT LAMA: \n' +
              'ONT BARU: \n' +
              'STB LAMA: \n' +
              'STB BARU: \n\n' +
              '<b>==TEKNISI==</b> \n'+
              'TEKNISI 1: \n' +
              'TEKNISI 2: \n' +
              'TEKNISI 3: \n\n' +
              '/help';
            sendText(id, notif);
          }
          break;
  
        //-----------------------------Command List STATUS--------------------------------//
        case "/statusPAREPARE":
          var sheetName = "PAREPARE"
          var args = chat_bot.split(" ").slice(1);
          sendText(id, searchByStatus(sheetName, args.join(" ")));
          break;
        case "/statusMAMUJU":
          var sheetName = "MAMUJU"
          var args = chat_bot.split(" ").slice(1);
          sendText(id, searchByStatus(sheetName, args.join(" ")));
          break;
        case "/statusSITOR":
          var sheetName = "SITOR"
          var args = chat_bot.split(" ").slice(1);
          sendText(id, searchByStatus(sheetName, args.join(" ")));
          break;
        case "/statusPALOPO":
          var sheetName = "PALOPO"
          var args = chat_bot.split(" ").slice(1);
          sendText(id, searchByStatus(sheetName, args.join(" ")));
          break;
        case "/statusHELP":
          let helpone = "<b>Halo !! Berikut Command Status yang dapat digunakan:</b>\n\n" +
            " /cekStatus Akses daftar STATUS \n" +
            " /statusPAREPARE Akses sheet PAREPARE\n" +
            " /statusMAMUJU Akses sheet MAMUJU\n" +
            " /statusSITOR Akses sheet SITOR\n" +
            " /statusPALOPO Akses sheet PALOPO\n\n" +
            " Template : /statusPAREPARE [status] \n\n" +
            " Back /help ";
          sendText(id, helpone);
          break;
        case "/cekStatus":
          let cekStatus = "<b>------Daftar status------</b>\n" +
            " ' OPEN \n" +
            " ' GAUL \n" +
            " ' GAMAS \n" +
            " ' GARANSI \n" +
            " ' REOPEN \n" +
            " 'OPEN SQM \n" +
            " * PREDATOR \n" +
            " * NO TIKET \n" +
            " 'GAMAS SQM \n" +
            " ################### \n" +
            " * TIKET PENDING \n" +
            " ^ MEDCARE \n" +
            " ^ SALTIK \n" +
            " ^ CLOSED \n" +
            " ^ CLOSED GAMAS \n" +
            " ^ CLOSED GAUL \n" +
            " ^ CLOSED GARANSI \n" +
            " ^CLOSED SQM \n\n" +
            "Back /statusHELP " +
            " Go to /help";
          sendText(id, cekStatus);
          break;

        //-----------------------------Command List STO--------------------------------//
        case "/stoPAREPARE":
          var sheetName = "PAREPARE"
          sendText(id, searchBySTO(sheetName, args));
          break;
        case "/stoMAMUJU":
          var sheetName = "MAMUJU"
          sendText(id, searchBySTO(sheetName, args));
          break;
        case "/stoSITOR":
          var sheetName = "SITOR"
          sendText(id, searchBySTO(sheetName, args));
          break;
        case "/stoPALOPO":
          var sheetName = "PALOPO"
          sendText(id, searchBySTO(sheetName, args));
          break;
        case "/stoHELP":
          let helptwo = "<b>Halo !! Berikut Command Status yang dapat digunakan:</b>\n\n" +
            " /cekSTO Akses daftar STO\n" +
            " /stoPAREPARE Akses sheet PAREPARE\n" +
            " /stoMAMUJU Akses sheet MAMUJU\n" +
            " /stoSITOR Akses sheet SITOR\n" +
            " /stoPALOPO Akses sheet PALOPO\n\n" +
            " Template : /stoPAREPARE [STO] \n\n" +
            " Back /help";
          sendText(id, helptwo);
          break;
        case "/cekSTO":
          let cekSTO = "<b>------KODE STO PAREPARE------</b> \n" +
            " BAR \n" +
            " PIN \n" +
            " PRE \n\n" +
            "<b>------KODE STO MAMUJU------</b> \n" +
            " MAJ \n" +
            " MAM \n" +
            " MMS \n" +
            " PLW \n" +
            " PKA \n" +
            " WON \n\n" +
            "<b>------KODE STO SITOR------</b> \n" +
            " AJN \n" +
            " CGI \n" +
            " ENR \n" +
            " MAK \n" +
            " RPN \n" +
            " RTP \n" +
            " SID \n" +
            " SIW \n" +
            " SKG \n" +
            " TTE \n" +
            " WTG \n\n" +
            "<b>------KODE STO PALOPO------</b> \n" +
            " MAS \n" +
            " MLL \n" +
            " PLP \n" +
            " SRK \n\n" +
            " Back /stoHELP \n" +
            " Go to /help";
          sendText(id, cekSTO);
          break;

        //-----------------------------Command List Detail--------------------------------//
        case "/detailPAREPARE":
          var sheetName = "PAREPARE"
          sendText(id, searchDetail(sheetName, args));
          break;
        case "/detailMAMUJU":
          var sheetName = "MAMUJU"
          sendText(id, searchDetail(sheetName, args));
          break;
        case "/detailSITOR":
          var sheetName = "SITOR"
          sendText(id, searchDetail(sheetName, args));
          break;
        case "/detailPALOPO":
          var sheetName = "PALOPO"
          sendText(id, searchDetail(sheetName, args));
          break;
        case "/detailHELP":
          let helpthree = "<b>Halo !! Berikut Command Detail yang dapat digunakan:</b>\n\n" +
            " /detailPAREPARE=Akses sheet PAREPARE\n" +
            " /detailMAMUJU Akses sheet MAMUJU\n" +
            " /detailSITOR Akses sheet SITOR\n" +
            " /detailPALOPO Akses sheet PALOPO\n\n" +
            " Template : /detailPAREPARE [NO USER] \n\n" +
            " Back /help";
          sendText(id, helpthree);
          break;

        //-----------------------------Exception Handling --------------------------------//
        default:
          let errormessage = "<b>Format yang dimasukkan salah</b> \n" +
            "Masukkan perintah sesuai petunjuk dan pastikan tag nama bot dihapus \n" +
            "cek /help ";
          sendText(id, errormessage);
      }
    } else {
      let error = "Perintah tidak tersedia!! /help";
      sendText(id, error);
    }
  }
}

//-----------------------------Current Time--------------------------------//
function getCurrentDateTime() {
  var today = new Date();
  var year = today.getFullYear();
  var month = String(today.getMonth() + 1).padStart(2, "0");
  var day = String(today.getDate()).padStart(2, "0");
  var hours = String(today.getHours()).padStart(2, "0");
  var minutes = String(today.getMinutes()).padStart(2, "0");
  var seconds = String(today.getSeconds()).padStart(2, "0");
  return year + "-" + month + "-" + day + " " + hours + ":" + minutes + ":" + seconds;
}

//-----------------------------Input data--------------------------------//
function writeToSpreadsheet(id, reportDate, inputer, namaPelanggan, contact, noUser, datel, customerType, alamatcust, laporanGangguan, sto) {
  
  if (sto === 'PRE' || sto === 'BAR' || sto === 'PIN') {
    sheetName = 'PAREPARE';
    var sheet = SpreadsheetApp.openById(SheetID).getSheetByName(sheetName);
  } else if (sto === 'MAJ' || sto === 'MAM' || sto === 'MMS' || sto === 'PLW' || sto === 'WON' || sto === 'PKA') {
    sheetName = 'MAMUJU';
    var sheet = SpreadsheetApp.openById(SheetID).getSheetByName(sheetName);
  } else if (sto === 'ENR' || sto === 'MAK' || sto === 'RPN' || sto === 'RTP' || sto === 'SKG' || sto === 'WTG' || sto === 'AJN' || sto === 'CGI' || sto === 'SID' || sto === 'SIW' || sto === 'TTE') {
    sheetName = 'SITOR';
    var sheet = SpreadsheetApp.openById(SheetID).getSheetByName(sheetName);
  } else if (sto === 'MAS' || sto === 'MLL' || sto === 'PLP' || sto === 'SRK') {
    sheetName = 'PALOPO';
    var sheet = SpreadsheetApp.openById(SheetID).getSheetByName(sheetName);
  } else {
    return sendText(id, "Laporan tidak berhasil di Input ! back to /help");
  }

    var data = sheet.getDataRange().getValues();
    var emptyRowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === "") {
        emptyRowIndex = i;
        break;
      }
    }

    if (emptyRowIndex !== -1) {
      var row = emptyRowIndex + 1;
      sheet.getRange(row, 1).setValue(reportDate);
      sheet.getRange(row, 8).setValue(namaPelanggan);
      sheet.getRange(row, 9).setValue(contact);
      sheet.getRange(row, 10).setValue(noUser);
      sheet.getRange(row, 11).setValue(datel);
      sheet.getRange(row, 12).setValue(customerType);
      sheet.getRange(row, 13).setValue(alamatcust);
      sheet.getRange(row, 16).setValue(laporanGangguan);
      sheet.getRange(row, 17).setValue(inputer);
      sheet.getRange(row, 18).setValue(sto);
    } else {
      var lastRow = data.length;
      var newRow = lastRow + 1;
      sheet.getRange(newRow, 1).setValue(reportDate);
      sheet.getRange(newRow, 8).setValue(namaPelanggan);
      sheet.getRange(newRow, 9).setValue(contact);
      sheet.getRange(newRow, 10).setValue(noUser);
      sheet.getRange(newRow, 11).setValue(datel);
      sheet.getRange(newRow, 12).setValue(customerType);
      sheet.getRange(newRow, 13).setValue(alamatcust);
      sheet.getRange(newRow, 16).setValue(laporanGangguan);
      sheet.getRange(newRow, 17).setValue(inputer);
      sheet.getRange(newRow, 18).setValue(sto);
    }
  sendText(id, "Laporan berhasil di Input ! back to /help");
}

function writeClose(groupName, id, reportDate, inputer, close, tiket, jenisGangguan, sektor, sto, user, odp, rfo, penyebabGangguan, dropcore, soc, roset, adaptor, sleeve, otp, prekso, heatshink, dll, ontLama, ontBaru, stbLama, stbBaru, teknisi1, teknisi2, teknisi3) {  
  if (id) {
    sheetName = 'REKAP';
    var sheet = SpreadsheetApp.openById(SheetID).getSheetByName(sheetName);
  } else {
    return sendText(id, "Laporan tidak berhasil di Input ! back to /help");
  }

    var data = sheet.getDataRange().getValues();
    var emptyRowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      if (data[i][1] === "") {
        emptyRowIndex = i;
        break;
      }
    }

    if (emptyRowIndex !== -1) {
      var row = emptyRowIndex + 1;
      sheet.getRange(row, 1).setValue(groupName);
      sheet.getRange(row, 2).setValue(reportDate);
      sheet.getRange(row, 3).setValue(inputer);
      sheet.getRange(row, 4).setValue(close);
      sheet.getRange(row, 5).setValue(tiket);
      sheet.getRange(row, 6).setValue(jenisGangguan);
      sheet.getRange(row, 7).setValue(sektor);
      sheet.getRange(row, 8).setValue(sto);
      sheet.getRange(row, 9).setValue(user);
      sheet.getRange(row, 10).setValue(odp);
      sheet.getRange(row, 11).setValue(rfo);
      sheet.getRange(row, 12).setValue(penyebabGangguan);
      sheet.getRange(row, 13).setValue(dropcore);
      sheet.getRange(row, 14).setValue(soc);
      sheet.getRange(row, 15).setValue(roset);
      sheet.getRange(row, 16).setValue(adaptor);
      sheet.getRange(row, 17).setValue(sleeve);
      sheet.getRange(row, 18).setValue(otp);
      sheet.getRange(row, 19).setValue(prekso);
      sheet.getRange(row, 20).setValue(heatshink);
      sheet.getRange(row, 21).setValue(dll);
      sheet.getRange(row, 22).setValue(ontLama);
      sheet.getRange(row, 23).setValue(ontBaru);
      sheet.getRange(row, 24).setValue(stbLama);
      sheet.getRange(row, 25).setValue(stbBaru);
      sheet.getRange(row, 26).setValue(teknisi1);
      sheet.getRange(row, 27).setValue(teknisi2);
      sheet.getRange(row, 28).setValue(teknisi3);
    } else {
      var lastRow = data.length;
      var newRow = lastRow + 1;
      sheet.getRange(newRow, 1).setValue(groupName);
      sheet.getRange(newRow, 2).setValue(reportDate);
      sheet.getRange(newRow, 3).setValue(inputer);
      sheet.getRange(newRow, 4).setValue(close);
      sheet.getRange(newRow, 5).setValue(tiket);
      sheet.getRange(newRow, 6).setValue(jenisGangguan);
      sheet.getRange(newRow, 7).setValue(sektor);
      sheet.getRange(newRow, 8).setValue(sto);
      sheet.getRange(newRow, 9).setValue(user);
      sheet.getRange(newRow, 10).setValue(odp);
      sheet.getRange(newRow, 11).setValue(rfo);
      sheet.getRange(newRow, 12).setValue(penyebabGangguan);
      sheet.getRange(newRow, 13).setValue(dropcore);
      sheet.getRange(newRow, 14).setValue(soc);
      sheet.getRange(newRow, 15).setValue(roset);
      sheet.getRange(newRow, 16).setValue(adaptor);
      sheet.getRange(newRow, 17).setValue(sleeve);
      sheet.getRange(newRow, 18).setValue(otp);
      sheet.getRange(newRow, 19).setValue(prekso);
      sheet.getRange(newRow, 20).setValue(heatshink);
      sheet.getRange(newRow, 21).setValue(dll);
      sheet.getRange(newRow, 22).setValue(ontLama);
      sheet.getRange(newRow, 23).setValue(ontBaru);
      sheet.getRange(newRow, 24).setValue(stbLama);
      sheet.getRange(newRow, 25).setValue(stbBaru);
      sheet.getRange(newRow, 26).setValue(teknisi1);
      sheet.getRange(newRow, 27).setValue(teknisi2);
      sheet.getRange(newRow, 28).setValue(teknisi3);
    }
    sendText(id, "Laporan berhasil di Input ! back to /help");
} 

//-----------------------------Nofication List--------------------------------//
const SHEETS_TO_MONITOR = ["PAREPARE", "MAMUJU", "SITOR", "PALOPO"];
const STATUS_COLUMN_INDEX = 5;
const NAME_COLUMN_INDEX = 8;
const ID_COLUMN_INDEX = 10;
const STO_COLUMN_INDEX = 18;

function onEdit(e) {
  var sheetName = e.source.getActiveSheet().getName();

  if (SHEETS_TO_MONITOR.includes(sheetName)) {
    var editedRange = e.range;
    var editedRow = editedRange.getRow();
    var editedCol = editedRange.getColumn();

    if (editedCol === STATUS_COLUMN_INDEX) {
      var newStatus = e.value;
      var name = e.source.getActiveSheet().getRange(editedRow, NAME_COLUMN_INDEX).getValue();
      var id = e.source.getActiveSheet().getRange(editedRow, ID_COLUMN_INDEX).getValue();
      var sto = e.source.getActiveSheet().getRange(editedRow, STO_COLUMN_INDEX).getValue();

      if (
        newStatus === "^ CLOSED" ||
        newStatus === "^ CLOSED GAMAS" ||
        newStatus === "^ CLOSED GAUL" ||
        newStatus === "^ CLOSED GARANSI" ||
        newStatus === "^CLOSED SQM"
      ) {
        if (name !== "") {
          var message = "<b>=====CLOSE REPORT=====</b> \n"+
                        "Status: " + newStatus + "\n";
          message += "Name: " + name + "\n";
          if (id) {
            message += "ID: " + id + "\n";
          }
          if (sto) {
            message += "STO: " + sto + "\n\n"+
                      "Wilayah: "+sheetName+"\n"+
                      "cek /detail"+sheetName+" [NO USER]"+"\n\n"+
                      "/detailHELP or /help ";
          }

          var groupUsernames = [-1001809863893,-967354257];
          for (var i = 0; i < groupUsernames.length; i++) {
          sendText(groupUsernames[i], message);
          }
        }
      }
    }
  }
}

//-------------------------------ChooseSheet--------------------------------//
function AmbilSheet(sheetName){
  var rangeName = sheetName + '!A2:AY';
  var rows = Sheets.Spreadsheets.Values.get(SheetID, rangeName).values;
  return rows;
}

function ReportSheet(){
  var rangeName = 'REPORT!CP8:CU31';
  var rows = Sheets.Spreadsheets.Values.get(SheetID, rangeName).values;
  return rows;
}

function TotalreportSheet(){
  var rangeName = 'REPORT!CP33:CU33';
  var rowsSum = Sheets.Spreadsheets.Values.get(SheetID, rangeName).values;
  return rowsSum;
}

//-----------------------------ReportingStat--------------------------------//
function sendReporting(){
  var report = ReportSheet();
  var sumReport = TotalreportSheet();
  const hasilPencarian = [];
  const total = [];
  var rowsTime = Sheets.Spreadsheets.Values.get(SheetID, 'REPORT!CP4').values;

  var th =  'ASSURANCE \n'+ rowsTime + '\n\n' +
            '<b>| STO | OPEN | CLOSED | COMPLY | NOT COMPLY | % |</b>\n' +
            '----------------------------------------------------------------------------------------\n' ;

  for (var row = 0; row < report.length; row++) {
      var infoReport = report[row][0] + ' ' + ' | ' + ' ' + report[row][1] + ' ' + ' | ' + ' ' + report[row][2] + ' ' +  ' | ' + ' ' + report[row][3] + ' ' + ' | ' + ' ' + report[row][4] + ' ' + ' | ' + ' ' + report[row][5];
      hasilPencarian.push(infoReport);
  }

  for (var row = 0; row < sumReport.length; row++) {
      var calReport = sumReport[row][0] + ' ' + ' | ' + ' ' + sumReport[row][1] + ' ' + ' | ' + ' ' + sumReport[row][2] + ' ' +  ' | ' + ' ' + sumReport[row][3] + ' ' + ' | ' + ' ' + sumReport[row][4] + ' ' + ' | ' + ' ' + sumReport[row][5];
      total.push(calReport);
  }

  var end = '\n-----------------------------------------------------------------------------------------\n' +
            total.join() + '\n\n' +
            " Back /help";
  
  if (hasilPencarian.length > 0) {
    var result = th + hasilPencarian.join("\n") + end
    return result;
  } else {
    var output = "<b>Reporting Tidak Ditemukan</b>\n\n";
    return output;
  }
}

function automateReport() {
  var id = updates.message.chat.id;
  sendText(id, sendReporting())
}

//-----------------------------searchByStatus--------------------------------//
function searchByStatus(sheetName, IDcek){
  var dataStatus = AmbilSheet(sheetName); 
  var num = 1;
  const hasilPencarian = [];

  for (var row = 0; row < dataStatus.length && hasilPencarian.length < 50; row++) {
    if(dataStatus[row][4] == IDcek) {
      var infoStatus = num++ + '.' + ' | ' + dataStatus[row][6] + ' | ' + dataStatus[row][7] + ' | ' + dataStatus[row][9] + ' | ' + dataStatus[row][17];
      hasilPencarian.push(infoStatus);
    }
  }

  var th =  '--------HASIL PENCARIAN STATUS '+'<b>'+IDcek+'</b>'+'-------- \n\n'+
            '<b> No. | NO TICKET | NAMA PELANGGAN | NO USER | STO </b>\n' +
            '----------------------------------------------------------------------------------------\n' ;

  var end = '\n-----------------------------------------------------------------------------------------\n' +
            '<b>Sheet yang diakses saat ini: </b>'+sheetName+'\n'+
            ' Template : /detail'+sheetName+' '+'[NO USER] \n\n' +
            ' Back /help atau /statusHELP';

  if (hasilPencarian.length > 0) {
    var result = th + hasilPencarian.join("\n\n") + end
    return result;
  } else {
    var output = "<b>Data STATUS Tidak Ditemukan</b>\n\n" +
                "Pastikan penulisan sesuai /cekStatus";
    return output;
  }
}

//-----------------------------searchBySTO--------------------------------//
function searchBySTO(sheetName, IDcek){
  var dataSTO = AmbilSheet(sheetName); 
  var num = 1;
  const hasilPencarian = [];

  for (var row = 0; row < dataSTO.length && hasilPencarian.length < 50; row++) {
    if(dataSTO[row][17] == IDcek) {
      var infoSTO = num++ + '.' + ' | ' + dataSTO[row][4] + ' | ' + dataSTO[row][6] + ' | ' + dataSTO[row][7] + ' | ' + dataSTO[row][9];
      hasilPencarian.push(infoSTO);
    }
  }

  var th =  '--------HASIL PENCARIAN STO '+'<b>'+IDcek+'</b>'+'-------- \n\n'+ 
            '<b> No. | STATUS | NO TICKET | NAMA PELANGGAN | NO USER </b>\n' +
            '-----------------------------------------------------------------------------------------\n' ;
             
  var end = '\n-----------------------------------------------------------------------------------------\n' +
            '<b>Sheet yang diakses saat ini: </b>'+sheetName+'\n'+
            ' Template : /detail'+sheetName+' '+'[NO USER] \n\n' +
            " Back /help atau /stoHELP";

  if (hasilPencarian.length > 0) {
    var result = th + hasilPencarian.join("\n\n") + end
    return result;
  } else {
    var output = "<b>Data STO Tidak Ditemukan</b>\n\n" +
                  "Pastikan STO dan penulisan sesuai /cekSTO";
    return output;
  }
}

//------------------------------search Detail Data--------------------------------//
function searchDetail(sheetName, IDcek){
  var dataDet = AmbilSheet(sheetName); 
  const hasilPencarian = [];
  
  var end = '\n-----------------------------------------------------------------------------------------\n' +
            "<b>Sheet yang diakses saat ini: </b>"+sheetName+"\n\n"+ 
            " Back /help atau /detailHELP";

  for (var row = 0; row < dataDet.length; row++) {
    if(dataDet[row][9] == IDcek) {
      var infodetail = "<b>-------- DETAIL "+dataDet[row][7]+" ---------</b>\n\n" +
                    "<b>Report Date : </b>"+ dataDet[row][0] + "\n" +
                    "<b>LIMIT 3 JAM : </b>"+ dataDet[row][1] + "\n" +
                    "<b>TTR 3 JAM : </b>"+ dataDet[row][2] + "\n" +
                    "<b>Last Update : </b>"+ dataDet[row][3] + "\n" +
                    "<b>STATUS : </b>"+ dataDet[row][4] + "\n" +
                    "<b>NO : </b>"+ dataDet[row][5] + "\n" +
                    "<b>NOMOR TIKET : </b>"+ dataDet[row][6] + "\n" +
                    "<b>NAMA PELANGGAN : </b>"+ dataDet[row][7] + "\n" +
                    "<b>CONTACT : </b>"+ dataDet[row][8] + "\n" +
                    "<b>NO USER : </b>"+ dataDet[row][9] + "\n" +
                    "<b>NO USER : </b>"+ dataDet[row][10] + "\n" +
                    "<b>CUSTOMER TYPE : </b>"+ dataDet[row][11] + "\n" +
                    "<b>KETERANGAN : </b>"+ dataDet[row][12] + "\n" +
                    "<b>A : </b>"+ dataDet[row][13] + "\n" +
                    "<b>F : </b>"+ dataDet[row][14] + "\n" +
                    "<b>Laporan Gangguan : </b>"+ dataDet[row][15] + "\n" +
                    "<b>INPUTER : </b>"+ dataDet[row][16] + "\n" +
                    "<b>STO : </b>"+ dataDet[row][17] + "\n" +
                    "<b>ODC : </b>"+ dataDet[row][18] + "\n" +
                    "<b>SEKTOR : </b>"+ dataDet[row][19] + "\n" +
                    "<b>DATEL : </b>"+ dataDet[row][20] + "\n" +
                    "<b>TTR 12 JAM : </b>"+ dataDet[row][21] + "\n" +
                    "<b>TTR 12 JAM SILVER : </b>"+ dataDet[row][22] + "\n" +
                    "<b>TTR 48 JAM : </b>"+ dataDet[row][23] + "\n" +
                    "<b>DURASI : </b>"+ dataDet[row][49] + "\n" +
                    "<b>WITEL : </b>"+ dataDet[row][50];
      hasilPencarian.push(infodetail);
    }
  }
  
  if (hasilPencarian.length > 0) {
    return hasilPencarian.join("\n\n") + end;
  } else {
    var output = "<b>Data detail Tidak Ditemukan</b>\n\n"+
                "Pastikan gunakan NO USER /detailHELP";
    return output;
  }
}

//-----------------------------End point--------------------------------//
function sendText(chatid,text,replymarkup){
  var data = {
    method: "post",
    payload: {
      method: "sendMessage",
      chat_id: String(chatid),
      text: text,
      parse_mode: "HTML",
      reply_markup: JSON.stringify(replymarkup)
    }
  };
  UrlFetchApp.fetch('https://api.telegram.org/bot' + token + '/', data);
} 
//aebiltaskari

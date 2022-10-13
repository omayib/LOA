function test(){
  var a = _readDosenByName(getSheet(getConfig().SHEET_NAME_DOSEN),"Arif Akbarul Huda, S.Si., M.Eng.");
  console.log(a);
}
function getSheet(sheetName){
  var spreadSheet = SpreadsheetApp.openById(getConfig().SHEET_ID);
  return spreadSheet.getSheetByName(sheetName);
}
function onSubmit(){
  var sheet = getSheet(getConfig().SHEET_NAME_MASTER);
  var lastRow = sheet.getLastRow();
  
  var nama_mahasiswa = sheet.getRange(lastRow, getConfig().COLUMN.NAMA_MAHASISWA).getValue();
  var judul = sheet.getRange(lastRow, getConfig().COLUMN.JUDUL).getValue();
  var target_publikasi = sheet.getRange(lastRow, getConfig().COLUMN.TARGET_PUBLIKASI).getValue();
  var nama_dosen = sheet.getRange(lastRow, getConfig().COLUMN.NAMA_DOSEN).getValue();
  
  var dosen = _readDosenByName(getSheet(getConfig().SHEET_NAME_DOSEN),nama_dosen).shift();
  var mhs = _readMahasiswaByName(getSheet(getConfig().SHEET_NAME_MHS),nama_mahasiswa).shift();
  var email_dosen = dosen.email
  var email_mahasiswa = mhs.email_mhs;
  var nim = mhs.nim;

  updateRowMaster(lastRow,nim,email_mahasiswa,email_dosen)

  var fileLoa = generatePdf(nama_mahasiswa, judul, nim,
          target_publikasi, nama_dosen);
  updateRowPenerbitan(nim,"DRAFT", fileLoa.getUrl());
  
  sentEmail(email_dosen,"Konfirmasi LOA","Confirmation",fileLoa,nim);
  
}
function updateForm(){
  // call your form and connect to the drop-down item
  var form = FormApp.openById(getConfig().FORM_ID);
  //1730434405
  var itemListDosen = form.getItemById(getConfig().LIST_DOSEN_ITEM_ID).asListItem();
  var sheetDosen = getSheet(getConfig().SHEET_NAME_DOSEN);
  var itemListMhs = form.getItemById(getConfig().LIST_MAHASISWA_ITEM_ID).asListItem();
  var sheetMhs = getSheet(getConfig().SHEET_NAME_MHS);

  // grab the values in the first column of the sheet - use 2 to skip header row
  var nameDosenValues = sheetDosen.getRange(2, 2, sheetDosen.getMaxRows() - 1).getValues();
  var nameMhsValues = sheetMhs.getRange(2, 2, sheetMhs.getMaxRows() - 1).getValues();
  console.log("mhs "+nameMhsValues);
  var dosenNames = [];
  var mhsNames = [];

  // convert the array ignoring empty cells
  for(var i = 0; i < nameDosenValues.length; i++)   
    if(nameDosenValues[i][0] != "")
      dosenNames[i] = nameDosenValues[i][0];
  itemListDosen.setChoiceValues(dosenNames);


  // convert the array ignoring empty cells
  for(var i = 0; i < nameMhsValues.length; i++)   
    if(nameMhsValues[i][0] != "")
      mhsNames[i] = nameMhsValues[i][0];
  itemListMhs.setChoiceValues(mhsNames);


}
function updateRowMaster(rowNumber,nim,email_mhs,email_dosen){
  console.log(rowNumber);
  console.log(nim);
  console.log(email_mhs);
  console.log(email_dosen);
  var sheet = getSheet(getConfig().SHEET_NAME_MASTER);
   sheet.getRange(rowNumber,getConfig().COLUMN.NIM_MAHASISWA).setValue(nim);
   sheet.getRange(rowNumber,getConfig().COLUMN.EMAIL_MAHASISWA).setValue(email_mhs);
   sheet.getRange(rowNumber,getConfig().COLUMN.EMAIL_DOSEN).setValue(email_dosen);
}
function updateRowPenerbitan(nim,status,fileUrl){
  var sheet = getSheet(getConfig().SHEET_NAME_MASTER);
   var data = _read(sheet,nim);
   var itemLoa = data.shift();
   sheet.getRange(itemLoa['row'],getConfig().COLUMN.STATUS).setValue(status);
   sheet.getRange(itemLoa['row'],getConfig().COLUMN.FILE_LOA).setValue(fileUrl);
}
function generatePdf(...args) {
  console.log(args);
  
  var tglPengajuan = Utilities.formatDate(new Date(),SpreadsheetApp.getActive().getSpreadsheetTimeZone(),"dd-MM-YYYY HH:mm:ss")

  var copiedFile = DriveApp.getFileById(getConfig().TEMPLATE_DOC_ID).makeCopy(),
      copiedId = copiedFile.getId(),
      copiedDoc = DocumentApp.openById(copiedId),
      copiedBody = copiedDoc.getActiveSection();
  
  copiedBody.replaceText('<< judul >>',args[1]);
  copiedBody.replaceText('<< target_publikasi >>',args[3]);
  copiedBody.replaceText('<< nama_mahasiswa >>',args[0]);
  copiedBody.replaceText('<< nim_mahasiswa >>',args[2]);
  copiedBody.replaceText('<< nama_dosen >>',args[4]);
  copiedBody.replaceText('<< tanggal >>',tglPengajuan);
  copiedDoc.saveAndClose();
  var newFile = DriveApp.getFolderById(getConfig().ARCHIVE_DIR_ID).createFile(copiedFile.getAs('application/pdf'));
  newFile.setName("LOA-"+args[2]+".pdf");
  copiedFile.setTrashed(true);

  return newFile;
}
function getIdFromUrl(url) { return url.match(/[-\w]{25,}/).shift(); }
function doGet(req){
  var nim = req.parameter.nim;
  var data = _read(getSheet(getConfig().SHEET_NAME_MASTER),nim);
  var itemLoa = data.shift();
  var file=DriveApp.getFileById(getIdFromUrl(itemLoa['file_LOA']));
  sentEmail(itemLoa['Email Mahasiswa'],
    "Selamat, LOA telah diterbitkan",
    "LOA-released",
    file,
    itemLoa['NIM']);
  sentEmail(itemLoa['Email Dosen'],
    "Selamat, LOA telah diterbitkan",
    "LOA-released",
    file,
    itemLoa['NIM']);
  updateRowPenerbitan(nim,"PUBLISHED",file.getUrl());
  
   var template = HtmlService.createTemplateFromFile('LOA-released');
   return template.evaluate().setSandboxMode(HtmlService.SandboxMode.IFRAME).setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
function sentEmail(email_to, subject, template_name, attachment, nim){
  if(MailApp.getRemainingDailyQuota() == 0){
    return;
  }
   var template = HtmlService.createTemplateFromFile(template_name);
   template.deploymentId = getConfig().DEPLOYMENT_ID;
   template.nim = nim;
    var message = template.evaluate().getContent();
    var options =  {
      htmlBody: message, 
      name: "Prodi Informatika",
      attachments: [attachment.getAs(MimeType.PDF)]
      };
  GmailApp.sendEmail(
        email_to, subject, '',options
      );
}
function _readDosenByName( sheet, name ) {
  var data         = sheet.getDataRange().getValues();
  var header       = data.shift();
  
  // Find All
  var result = data.map(function( row, indx ) {
    var reduced = header.reduce( function(accumulator, currentValue, currentIndex) {
      accumulator[ currentValue ] = row[ currentIndex ];
      return accumulator;
    }, {});

    reduced.row = indx + 2;
    return reduced;
    
  });
  
  // Filter if id is provided
  if( name ) {
    var filtered = result.filter( function( record ) {
      if ( record['Nama'].toUpperCase() === name.toUpperCase() ) {
        return true;
      } else {
        return false;
      }
    });
    
    return filtered;
  } 
  
  return result; 
}
function _readMahasiswaByName( sheet, name ) {
  var data         = sheet.getDataRange().getValues();
  var header       = data.shift();
  
  // Find All
  var result = data.map(function( row, indx ) {
    var reduced = header.reduce( function(accumulator, currentValue, currentIndex) {
      accumulator[ currentValue ] = row[ currentIndex ];
      return accumulator;
    }, {});

    reduced.row = indx + 2;
    return reduced;
    
  });
  
  // Filter if id is provided
  if( name ) {
    var filtered = result.filter( function( record ) {
      if ( record['nama_lengkap'].toUpperCase() === name.toUpperCase() ) {
        return true;
      } else {
        return false;
      }
    });
    
    return filtered;
  } 
  
  return result; 
}
function _read( sheet, nim ) {
  var data         = sheet.getDataRange().getValues();
  var header       = data.shift();
  
  // Find All
  var result = data.map(function( row, indx ) {
    var reduced = header.reduce( function(accumulator, currentValue, currentIndex) {
      accumulator[ currentValue ] = row[ currentIndex ];
      return accumulator;
    }, {});

    reduced.row = indx + 2;
    return reduced;
    
  });
  
  // Filter if id is provided
  if( nim ) {
    var filtered = result.filter( function( record ) {
      if ( record.NIM.toUpperCase() === nim.toUpperCase() ) {
        return true;
      } else {
        return false;
      }
    });
    
    return filtered;
  } 
  
  return result; 
}
function response() {
   return {
      json: function(data) {
         return ContentService
            .createTextOutput(JSON.stringify(data))
            .setMimeType(ContentService.MimeType.JSON);
      }
   }
}
// Google Apps Script – tempel di script.google.com (terikat ke Google Sheet)
const FOLDER_NAME = 'Foto Laporan Karhutla Gedong';
const SHEET = 'Laporan';

function sh(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET) || ss.insertSheet(SHEET);
  if (!s.getLastRow()) s.appendRow(['Waktu','Nama','HP','Kondisi','Keterangan','Latitude','Longitude','Link Foto','Link Peta']);
  return s;
}
function doPost(e){
  try{
    const d = JSON.parse(e.postData.contents);
    const it = DriveApp.getFoldersByName(FOLDER_NAME);
    const folder = it.hasNext() ? it.next() : DriveApp.createFolder(FOLDER_NAME);
    const blob = Utilities.newBlob(Utilities.base64Decode(d.foto.split(',')[1]), 'image/jpeg', 'lapor_'+Date.now()+'.jpg');
    const f = folder.createFile(blob); f.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    sh().appendRow([new Date(), d.nama, "'"+d.hp, d.kondisi, d.ket, d.lat, d.lng, f.getUrl(), 'https://maps.google.com/?q='+d.lat+','+d.lng]);
    return out({ok:true});
  }catch(err){ return out({ok:false,error:String(err)}); }
}
function doGet(){
  const v = sh().getDataRange().getValues().slice(1).reverse().slice(0,15);
  return out({rows:v.map(r=>({waktu:r[0],kondisi:r[3],lat:r[5],lng:r[6]}))}); // data pribadi (nama/HP/foto) tidak ditampilkan
}
function out(o){ return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

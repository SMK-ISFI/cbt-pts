/** Soal dibaca dari satu file soal.json privat di Google Drive. */
const SETTINGS = {
  questionFileId: "TEMPEL_ID_FILE_SOAL_JSON",
  // ID spreadsheet disimpan otomatis oleh fungsi setup.
  attemptsSheet: 'Sesi',
  resultsSheet: 'Nilai'
};
function setup() {
  const BANK = readBank_();
  const ss=SpreadsheetApp.getActiveSpreadsheet();
  if(!ss)throw new Error('Buka Apps Script melalui menu Ekstensi pada Google Sheet.');
  PropertiesService.getScriptProperties().setProperty('SHEET_ID',ss.getId());
  const roster=table_(ss,'Peserta',['ID','Nama','Kelas']);
  if(roster.getLastRow()===1)BANK.students.forEach(s=>roster.appendRow([s.id,s.name,s.className]));
  table_(ss,SETTINGS.attemptsSheet,['Attempt ID','Exam ID','Student ID','Token','Mulai (ms)','Deadline (ms)','Snapshot JSON']);
  table_(ss,SETTINGS.resultsSheet,['Attempt ID','Exam ID','Student ID','Nama','Kelas','Mulai','Selesai klien','Diterima server','Alasan','Benar','Salah','Kosong','Jumlah soal','Nilai','Jawaban JSON','Terlambat diterima']);
}
function table_(ss,name,headers){let sh=ss.getSheetByName(name);if(!sh){sh=ss.insertSheet(name);sh.appendRow(headers);sh.setFrozenRows(1);sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#245a49').setFontColor('#ffffff');}return sh;}
function doGet(){return json_({ok:true,message:'CBT API aktif. Gunakan POST dari aplikasi CBT.'});}
function doPost(e){
  const lock=LockService.getScriptLock();
  try {
    const p=JSON.parse(e.postData.contents);
    if(p.action==='bootstrap'){const BANK=readBank_();return json_({ok:true,exam:BANK.exam,students:students_(),questionCount:BANK.questions.length,serverNow:Date.now()});}
    lock.waitLock(20000);
    const ss=spreadsheet_(), attempts=ss.getSheetByName(SETTINGS.attemptsSheet), results=ss.getSheetByName(SETTINGS.resultsSheet);
    if(!attempts||!results)throw new Error('Guru perlu menjalankan fungsi setup terlebih dahulu.');
    if(p.action==='start')return json_(start_(p,attempts));
    const row=find_(attempts,p.attemptId);
    if(!row||typeof p.token!=='string'||row[3]!==p.token)throw new Error('Sesi tidak ditemukan atau token tidak sesuai. Hubungi guru.');
    const previous=find_(results,p.attemptId);
    if(p.action==='resume')return json_({ok:true,deadline:Number(row[5]),serverNow:Date.now(),result:previous?result_(previous):null});
    if(p.action!=='submit')throw new Error('Aksi tidak dikenal.');
    // Retry selalu mengembalikan hasil yang sama, tidak menulis baris kedua.
    if(previous)return json_({ok:true,result:result_(previous),serverNow:Date.now()});
    const snap=JSON.parse(row[6]), time=Date.now(), deadline=Number(row[5]);
    if(time<deadline-snap.exam.finishWindowMinutes*60000)throw new Error('Ujian belum memasuki 10 menit terakhir.');
    if(!p.answers||typeof p.answers!=='object'||Array.isArray(p.answers))throw new Error('Format jawaban tidak valid.');
    let correct=0,answered=0;const clean={};
    snap.questions.forEach(q=>{const a=p.answers[q.id];if(a===undefined||a===null||a==='')return;if(!Object.prototype.hasOwnProperty.call(q.options,a))throw new Error('Pilihan jawaban tidak valid.');clean[q.id]=a;answered++;if(a===q.answer)correct++;});
    const total=snap.questions.length, score=Math.round(correct/total*10000)/100;
    const clientEnd=Number.isFinite(p.finishedAt)?Math.max(Number(row[4]),Math.min(p.finishedAt,deadline)):deadline;
    const reason=time>=deadline?'timeout':(p.reason==='manual'?'manual':'timeout');
    const out=[p.attemptId,snap.exam.id,snap.student.id,safe_(snap.student.name),safe_(snap.student.className),new Date(Number(row[4])),new Date(clientEnd),new Date(time),reason,correct,answered-correct,total-answered,total,score,JSON.stringify(clean),time>deadline+60000?'YA':'TIDAK'];
    results.appendRow(out);SpreadsheetApp.flush();
    return json_({ok:true,result:result_(out),serverNow:time});
  }catch(error){return json_({ok:false,error:String(error.message||error)});}
  finally{if(lock.hasLock())lock.releaseLock();}
}
function start_(p,sh){
  const BANK=readBank_();
  if(typeof p.attemptId!=='string'||!/^[a-zA-Z0-9-]{20,60}$/.test(p.attemptId))throw new Error('ID sesi tidak valid.');
  if(p.examId!==BANK.exam.id)throw new Error('Konfigurasi ujian berubah. Muat ulang halaman.');
  const student=students_().find(s=>s.id===p.studentId);if(!student)throw new Error('Peserta tidak terdaftar.');
  const old=find_(sh,p.attemptId);
  if(old){if(old[1]!==BANK.exam.id||old[2]!==student.id)throw new Error('ID sesi sudah dipakai.');return session_(old);}
  const rows=sh.getDataRange().getValues().slice(1);
  if(rows.some(r=>r[1]===BANK.exam.id&&r[2]===student.id))throw new Error('Nama ini sudah memiliki sesi ujian. Gunakan browser asal atau minta guru memulihkan sesi.');
  const start=Date.now(), snapshot={exam:BANK.exam,student,questions:BANK.questions};
  const row=[p.attemptId,BANK.exam.id,student.id,Utilities.getUuid(),start,start+BANK.exam.durationMinutes*60000,JSON.stringify(snapshot)];
  if(row[6].length>45000)throw new Error('Bank soal terlalu besar untuk versi sederhana ini. Kurangi panjang/jumlah soal.');
  sh.appendRow(row);SpreadsheetApp.flush();return session_(row);
}
function session_(row){const s=JSON.parse(row[6]);return {ok:true,exam:s.exam,student:s.student,questions:s.questions.map(q=>{const {answer,explanation,...publicQuestion}=q;return publicQuestion;}),token:row[3],startedAt:Number(row[4]),deadline:Number(row[5]),serverNow:Date.now()};}
function find_(sheet,id){if(typeof id!=='string')return null;const rows=sheet.getDataRange().getValues();for(let i=1;i<rows.length;i++)if(rows[i][0]===id)return rows[i];return null;}
function result_(r){return {correct:Number(r[9]),wrong:Number(r[10]),blank:Number(r[11]),total:Number(r[12]),score:Number(r[13])};}
function safe_(s){s=String(s);return /^[=+@\-]/.test(s)?"'"+s:s;}
function json_(value){return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);}

function spreadsheet_(){
  const id=PropertiesService.getScriptProperties().getProperty('SHEET_ID');
  if(!id)throw new Error('Jalankan setup sekali melalui editor Apps Script.');
  return SpreadsheetApp.openById(id);
}
function students_(){
  const sheet=spreadsheet_().getSheetByName('Peserta');
  if(!sheet)throw new Error('Tab Peserta belum ada. Jalankan setup.');
  const rows=sheet.getDataRange().getDisplayValues().slice(1).filter(r=>r[0]&&r[1]);
  const students=rows.map(r=>({id:r[0].trim(),name:r[1].trim(),className:r[2].trim()}));
  if(!students.length)throw new Error('Isi setidaknya satu nama pada tab Peserta.');
  if(new Set(students.map(s=>s.id)).size!==students.length)throw new Error('ID pada tab Peserta harus unik.');
  return students;
}

// Bank soal hanya ada di JSON, tidak ditulis di kode program.
function readBank_(){
  if(SETTINGS.questionFileId === 'TEMPEL_ID_FILE_SOAL_JSON')throw new Error('Isi questionFileId dengan ID soal.json dari Google Drive.');
  const file=DriveApp.getFileById(SETTINGS.questionFileId);
  const bank=JSON.parse(file.getBlob().getDataAsString('UTF-8').replace(/^\uFEFF/,''));
  if(!bank.exam || !bank.exam.id || !Array.isArray(bank.questions) || !bank.questions.length)throw new Error('Struktur soal.json tidak valid.');
  const d=bank.exam.durationMinutes, w=bank.exam.finishWindowMinutes;
  if(!Number.isFinite(d)||!Number.isFinite(w)||d<=0||w<=0||w>d)throw new Error('Durasi/jendela akhir pada JSON tidak valid.');
  if(new Set(bank.questions.map(q=>q.id)).size!==bank.questions.length)throw new Error('ID soal harus unik.');
  bank.questions.forEach(q=>{if(!q.id||typeof q.text!=='string'||!q.options||!Object.prototype.hasOwnProperty.call(q.options,q.answer))throw new Error('Soal atau kunci tidak valid: '+q.id);});
  if(!Array.isArray(bank.students))bank.students=[];
  return bank;
}

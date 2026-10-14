const vm=require('node:vm'),fs=require('node:fs'),assert=require('node:assert/strict'),test=require('node:test');
let clock=1000000000000;
class FakeDate extends Date { constructor(...args){super(...(args.length?args:[clock]));} static now(){return clock;} }
const sheets=new Map();
class Sheet {
 constructor(){this.rows=[];}
 appendRow(r){this.rows.push([...r]);}
 getLastRow(){return this.rows.length;} getDataRange(){return {getValues:()=>this.rows,getDisplayValues:()=>this.rows.map(r=>r.map(String))};}
 setFrozenRows(){} getRange(){return {setFontWeight(){return this;},setBackground(){return this;},setFontColor(){return this;}};}
}
const ss={getId:()=> 'test-sheet',getSheetByName:n=>sheets.get(n),insertSheet:n=>{const s=new Sheet();sheets.set(n,s);return s;}};
let held=false;
const ctx=vm.createContext({DriveApp:{getFileById:()=>({getBlob:()=>({getDataAsString:()=>fs.readFileSync('guru/soal.json','utf8')})})},Date:FakeDate,PropertiesService:{getScriptProperties:()=>({setProperty(){},getProperty:()=> 'test-sheet'})},SpreadsheetApp:{getActiveSpreadsheet:()=>ss,openById:()=>ss,flush(){}},Utilities:{getUuid:()=> 'server-token'},LockService:{getScriptLock:()=>({waitLock(){held=true;},hasLock:()=>held,releaseLock(){held=false;}})},ContentService:{MimeType:{JSON:'json'},createTextOutput:text=>({setMimeType:()=>text})}});
vm.runInContext(fs.readFileSync('guru/Code.gs','utf8').replace('TEMPEL_ID_FILE_SOAL_JSON','test-json-id'),ctx);
vm.runInContext('setup()',ctx);
const api=p=>JSON.parse(ctx.doPost({postData:{contents:JSON.stringify(p)}}));
const attemptId='test-attempt-0000000000001'; let session;
test('bootstrap dan sesi tidak membocorkan kunci jawaban',()=>{
 const boot=api({action:'bootstrap'}); assert.equal(boot.ok,true); assert.equal(boot.questionCount,25); assert.equal(boot.questions,undefined);
 session=api({action:'start',examId:'php-sdlc-25-v3',studentId:'S001',attemptId});
 assert.equal(session.ok,true); assert.equal(session.questions.some(q=>Object.hasOwn(q,'answer')||Object.hasOwn(q,'explanation')),false);
});
test('retry start idempotent dan peserta tidak dapat membuat sesi kedua',()=>{
 assert.equal(api({action:'start',examId:'php-sdlc-25-v3',studentId:'S001',attemptId}).token,session.token);
 assert.equal(sheets.get('Sesi').rows.length,2);
 assert.equal(api({action:'start',examId:'php-sdlc-25-v3',studentId:'S001',attemptId:attemptId+'2'}).ok,false);
});
test('token salah tidak dinilai; submit sah terlalu dini ditolak',()=>{
 const bad=api({action:'submit',attemptId,token:'wrong',answers:{}});
 assert.equal(bad.receipt.status,'needs_review');
 assert.equal(sheets.get('Nilai').rows.length,1);
 assert.equal(api({action:'submit',attemptId,token:session.token,answers:{q01:'C'}}).ok,false);
});
test('server menghitung nilai sendiri dan retry tidak menggandakan baris',()=>{
 clock=session.deadline-600000;
 const payload={action:'submit',attemptId,token:session.token,answers:{q01:'B',q02:'E'},reason:'manual',score:100,finishedAt:clock};
 const saved=api(payload); assert.equal(saved.ok,true); assert.equal(saved.result.score,4);
 const retry=api({...payload,answers:{}});assert.deepEqual(retry.result,saved.result);
 assert.equal(sheets.get('Nilai').rows.length,2);
 const resume=api({action:'resume',attemptId,token:session.token}); assert.equal(resume.result.score,4);
});
test('hasil yang terlambat karena offline tetap diterima dan ditandai',()=>{
 const second=api({action:'start',examId:'php-sdlc-25-v3',studentId:'S002',attemptId:attemptId+'3'});
 clock=second.deadline+120000;
 const saved=api({action:'submit',attemptId:attemptId+'3',token:second.token,answers:{q01:'C'},reason:'timeout',finishedAt:second.deadline});
 assert.equal(saved.ok,true);assert.equal(sheets.get('Nilai').rows[2][15],'YA');
});

test('baris Sesi terhapus: pulih dari cadangan dan dinilai tepat satu kali',()=>{
 const id='recovery-attempt-000000000001';
 const started=api({action:'start',examId:'php-sdlc-25-v3',studentId:'S003',attemptId:id});assert.equal(started.ok,true);
 assert.ok(sheets.get('CadanganSesi').rows.find(r=>r[0]===id));
 sheets.get('Sesi').rows=sheets.get('Sesi').rows.filter(r=>r[0]!==id);
 clock=started.deadline-600000;
 const request={action:'submit',attemptId:id,token:started.token,answers:{q01:'B'},finishedAt:clock,reason:'manual'};
 const result=api(request);assert.equal(result.ok,true);assert.equal(result.result.score,4);
 assert.ok(sheets.get('Sesi').rows.find(r=>r[0]===id));
 assert.equal(api(request).result.score,4);
 assert.equal(sheets.get('Nilai').rows.filter(r=>r[0]===id).length,1);
});
test('sesi dan cadangan hilang: jawaban lama diterima tanpa penilaian palsu',()=>{
 const id='legacy-attempt-0000000000001';
 const started=api({action:'start',examId:'php-sdlc-25-v3',studentId:'S004',attemptId:id});
 sheets.get('Sesi').rows=sheets.get('Sesi').rows.filter(r=>r[0]!==id);
 sheets.get('CadanganSesi').rows=sheets.get('CadanganSesi').rows.filter(r=>r[0]!==id);
 const resume=api({action:'resume',attemptId:id,token:started.token});assert.equal(resume.needsReview,true);
 const request={action:'submit',attemptId:id,token:started.token,answers:{q01:'B'},recovery:{examId:'php-sdlc-25-v3',student:{id:'S004',name:'Nama siswa',className:'XI PPLG 1'},questions:[{id:'q01',answer:'B'}]}};
 const result=api(request);assert.equal(result.ok,true);assert.equal(result.receipt.status,'needs_review');assert.equal(result.result,undefined);
 assert.equal(api(request).receipt.status,'needs_review');
 const rows=sheets.get('PerluVerifikasi').rows.filter(r=>r[0]===id);assert.equal(rows.length,1);assert.equal(JSON.parse(rows[0][7]).q01,'B');
 assert.equal(sheets.get('Nilai').rows.filter(r=>r[0]===id).length,0);
});
test('token sesi berubah di tab utama: cadangan asli tetap dapat dipakai',()=>{
 const id='changed-token-attempt-000001';
 const started=api({action:'start',examId:'php-sdlc-25-v3',studentId:'S005',attemptId:id});
 const row=sheets.get('Sesi').rows.find(r=>r[0]===id);row[3]='edited-token';
 clock=started.deadline;
 const result=api({action:'submit',attemptId:id,token:started.token,answers:{q01:'B'},reason:'timeout'});
 assert.equal(result.ok,true);assert.equal(result.result.score,4);
});

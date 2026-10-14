const vm=require('node:vm'),fs=require('node:fs'),assert=require('node:assert/strict'),test=require('node:test');
let clock=1000000000000;
class FakeDate extends Date { constructor(...args){super(...(args.length?args:[clock]));} static now(){return clock;} }
const sheets=new Map();
class Sheet {
 constructor(){this.rows=[];}
 appendRow(r){this.rows.push(r);}
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
test('server menolak token salah dan submit dini',()=>{
 assert.equal(api({action:'submit',attemptId,token:'wrong',answers:{}}).ok,false);
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

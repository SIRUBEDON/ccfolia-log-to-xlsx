const test=require('node:test');
const assert=require('node:assert/strict');
const {openSite}=require('./helpers.cjs');
test('CSV cells neutralize formula prefixes while preserving quotes, commas and line breaks',async()=>{
 const {dom,w}=await openSite();
 for(const value of ['=1+1','+SUM(A1)','-1+1','@SUM(A1)',' \t=1+1','\ttext','\rtext','\ntext','＝1+1','＋1','－1','＠SUM(A1)']) {
  assert.ok(w.serializeCsvCell(value).startsWith('"\''),JSON.stringify(value));
 }
 assert.equal(w.serializeCsvCell('あ,"い"\nう'),'"あ,""い""\nう"');
 assert.equal(w.serializeCsvCell(0),'"0"');
 assert.equal(w.serializeCsvCell('通常の発言'),'"通常の発言"');
 dom.window.close();
});
test('CSV export protects speaker, tab heading and message cells',async()=>{
 const {dom,w}=await openSite();
 let blob;w.URL.createObjectURL=b=>(blob=b,'blob:https://site.test/test');
 w.eval('parsedRows = [["=1+1", "=2+2", "=3+3", "#000000"]]');
 w.displayTabSelection();
 w.document.querySelector('input[value="csv"]').checked=true;
 await w.convertFile();
 const reader=new w.FileReader();
 const csv=await new Promise((resolve,reject)=>{reader.onload=()=>resolve(reader.result);reader.onerror=reject;reader.readAsText(blob);});
 assert.ok(csv.includes('"\'=2+2"\r\n'));
 assert.ok(csv.includes('"\'=1+1","\'=2+2","\'=3+3"'));
 dom.window.close();
});

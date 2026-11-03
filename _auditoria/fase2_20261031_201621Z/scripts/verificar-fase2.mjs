import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import assert from 'node:assert/strict';
import {fileURLToPath} from 'node:url';
import {parseProcedures,stripStringsAndComments,runAudit} from './auditar-vba.mjs';
const dir=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const sha=b=>crypto.createHash('sha256').update(b).digest('hex');
const decode=b=>{try{return new TextDecoder('utf-8',{fatal:true}).decode(b);}catch{return new TextDecoder('windows-1252').decode(b);}};
const sig=p=>p.body.slice(0,p.body.findIndex(l=>!l.trimEnd().endsWith('_'))+1).join('\n').trim();
export async function verificarFase2(){
 const m=JSON.parse(await fs.readFile(path.join(dir,'manifesto-inicial-fase2.json'),'utf8'));
 const old=JSON.parse(await fs.readFile(path.join(m.previous,'manifesto-original.json'),'utf8'));
 const checks=[];
 const check=async(name,fn)=>{try{checks.push({name,status:'PASS',details:await fn()});}catch(e){checks.push({name,status:'FAIL',error:e.message});}};
 const originalSources=old.files.filter(f=>/\.(bas|frm)$/i.test(f.relative));
 const before=new Map(),after=new Map();
 for(const f of originalSources){
  before.set(f.relative,decode(await fs.readFile(path.join(m.previous,'desenvolvimento',f.relative))));
  after.set(f.relative,decode(await fs.readFile(path.join(m.dev2,f.relative))));
 }
 await check('39 originais e backups da fase 1 intactos',async()=>{
  for(const f of old.files){
   assert.equal(sha(await fs.readFile(path.join(old.root,f.relative))),f.sha256);
   assert.equal(sha(await fs.readFile(path.join(old.backup,f.relative))),f.sha256);
  } return old.files.length;
 });
 await check('155 arquivos da execucao anterior intactos',async()=>{
  for(const f of m.previousInventory)assert.equal(sha(await fs.readFile(path.join(m.previous,f.relative))),f.sha256);
  return m.previousInventory.length;
 });
 await check('46 backups novos conferidos',async()=>{
  for(const f of m.backups)assert.equal(sha(await fs.readFile(path.join(m.backup2,f.kind,f.relative))),f.sha256);
  return m.backups.length;
 });
 await check('Cinco pares FRM/FRX e XLSM preservados',async()=>{
  for(const f of old.files.filter(f=>/\.(frx|xlsm)$/i.test(f.relative))) assert.equal(sha(await fs.readFile(path.join(m.dev2,f.relative))),f.sha256);
  for(const [f,b] of before) if(/\.frm$/i.test(f))assert.equal(after.get(f).split(/Option Explicit/i)[0],b.split(/Option Explicit/i)[0]);
  return 'Designers e FRX identicos; binario XLSM NAO incorpora revisao.';
 });
 await check('647 assinaturas exportadas existentes preservadas',async()=>{
  let count=0;
  for(const [f,b]of before){
   const revised=parseProcedures(after.get(f));
   for(const p of parseProcedures(b)){
    const found=revised.find(q=>q.name===p.name&&q.kind===p.kind);assert.ok(found,f+':'+p.name);assert.equal(sig(found),sig(p));count++;
   }
   assert.equal(after.get(f).match(/Attribute VB_Name = "[^"]+"/i)[0],b.match(/Attribute VB_Name = "[^"]+"/i)[0]);
  } assert.equal(count,647);return count;
 });
 await check('Codificacao e CRLF existentes mantidos',async()=>{
  for(const[f,b]of before){
   const raw=await fs.readFile(path.join(m.dev2,f));
   const oldRaw=await fs.readFile(path.join(m.previous,'desenvolvimento',f));
   const utf8=v=>{try{new TextDecoder('utf-8',{fatal:true}).decode(v);return true;}catch{return false;}};
   if(utf8(oldRaw)){assert.equal(sha(Buffer.from(after.get(f),'utf8')),sha(raw),f);}else{
    const decoder=new TextDecoder('windows-1252');const reverse=new Map();
    for(let k=0;k<256;k++)reverse.set(decoder.decode(Uint8Array.of(k)),k);
    const encoded=Buffer.from(Array.from(after.get(f),c=>{assert.ok(reverse.has(c),f+':caractere incompativel');return reverse.get(c);}));
    assert.equal(sha(encoded),sha(raw),f);
   }
   if(!/(?<!\r)\n/.test(b))assert.ok(!/(?<!\r)\n/.test(after.get(f)),f);
  }return before.size;
 });
 await check('Sete comandos Delete preservados literalmente',async()=>{
  const deletes=t=>t.split(/\r?\n/).filter(l=>/\.Delete\b/i.test(stripStringsAndComments(l))).map(l=>l.trim());
  let n=0;
  for(const[f,b]of before){assert.deepEqual(deletes(after.get(f)),deletes(b));n+=deletes(b).length;}
  assert.equal(n,7);return n;
 });
 await check('Rotinas fiscais especializadas identicas',async()=>{
  const f=[...before.keys()].find(k=>k.endsWith('modValidacaoXML.bas'));
  const now=parseProcedures(after.get(f));
  const records=[];
  for(const p of parseProcedures(before.get(f)).filter(p=>/CFOP|CST|Benef|ICMS|PIS|COFINS/i.test(p.name))){
   assert.equal(now.find(q=>q.name===p.name).body.join('\n'),p.body.join('\n'));records.push(p.name);
  }assert.ok(records.length>0);return records;
 });
 await check('A01 A02 A03 A04 A05 A06 A07 permanecem nos fontes',async()=>{
  const texts=[...after.values()].join('\n');
  const saida=[...after.entries()].find(([f])=>f.endsWith('Cad_NFE_Saida.frm'))[1];
  const gravar=parseProcedures(saida).find(p=>p.name==='CommandButton_Gravar_Click').body.join('\n');
  assert.ok(!/^\s*celula\.Errors/im.test(gravar));
  assert.ok(/\.Cells\(linhaDemo, 12\)\.Errors/i.test(gravar));
  assert.ok(!/^\s*(?:Me\.)?Frame_ImpostoFilial\./im.test(saida));
  assert.ok(/NumberFormat = "@"\r?\n\s*.*= CStr\(TextBoxChaveDeAcesso/i.test(texts));
  assert.ok(/ListIndex (?:< 0|= -1)/i.test(texts));
  for(const[f,t]of after)assert.ok(/^Option Explicit/im.test(t),f);
  const ribbon=[...after.entries()].find(([f])=>f.endsWith('MóduloOfficeRibbonEditor.bas'))[1];
  assert.ok(/ThisWorkbook.Save/i.test(ribbon));assert.ok(!/ActiveWorkbook.Save/i.test(ribbon));
  return 'Presenca estatica; suite anterior valida filtros por 35 cenarios JS, nao executa VBA.';
 });
 await check('Novos auxiliares e prevalidacoes antes das escritas',async()=>{
  let n=0;
  for(const[f,t]of after)if(/Cad_NFE(?:_Saida)?\.frm$/.test(f)){
   for(const p of parseProcedures(t).filter(p=>/^CommandButton_Gravar(?:Filial)?_Click$/.test(p.name))){
    const body=p.body.join('\n');
    assert.ok(body.indexOf('ValidarConversoesGravacao')<body.search(/ActiveCell\.value\s*=/i));
    assert.ok(body.indexOf('operacao.RegistrarDestino')<body.search(/ActiveCell\.value\s*=/i));
    assert.ok(body.includes('operacao.Recuperar(diagnostico)'));assert.ok(body.includes('operacao.Finalizar'));
    if(body.includes('.Delete'))assert.ok(body.indexOf('operacao.MarcarBaixaIniciada')<body.indexOf('.Delete'));
    if(body.includes('ThisWorkbook.Save'))assert.ok(body.indexOf('operacao.MarcarSalvamentoIniciado')<body.indexOf('ThisWorkbook.Save'));
    n++;
   }
  }assert.equal(n,4);return n;
 });
 await check('Sem novas aspas VBA abertas nos arquivos alterados',async()=>{
  let n=0;
  for(const folder of ['MDL','Form'])for(const f of await fs.readdir(path.join(m.dev2,folder))){
   if(!/\.(bas|cls|frm)$/i.test(f))continue;
   const t=decode(await fs.readFile(path.join(m.dev2,folder,f)));
   for(const[lineNo,l]of t.split(/\r?\n/).entries()){
    let quoted=false;
    for(let i=0;i<l.length;i++){
     if(l[i]==='"'){if(quoted&&l[i+1]==='"'){i++;continue;}quoted=!quoted;}
     if(!quoted&&l[i]==="'")break;
    }
    assert.ok(!quoted,f+':'+(lineNo+1));
   }n++;
  }return n+' fontes; verificacao lexical nao substitui compilacao.';
 });
 const audit=await runAudit(m.dev2,path.join(dir,'evidencias','estatica-final'));
 await check('Inventario final sem duplicidade publica global',async()=>{
  assert.equal(audit.files.length,34);assert.equal(audit.procedureCount,665);assert.equal(audit.publicDuplicates.length,0);
  return {files:34,exportedProcedures:665,existingEmbeddedWorkbookComponents:50,prospectiveComponentsWithoutTestModules:52,prospectiveTotalProcedures:666,unqualifiedSheets:audit.hits.unqualifiedSheets?.length,activeWorkbook:audit.hits.activeWorkbook?.length};
 });
 const result={timestamp:new Date().toISOString(),kind:'Analise estatica/JavaScript; NAO e compilacao VBA',checks,pass:checks.filter(t=>t.status==='PASS').length,fail:checks.filter(t=>t.status==='FAIL').length};
 await fs.writeFile(path.join(dir,'testes-estaticos-fase2.json'),JSON.stringify(result,null,2));
 return result;
}
if(globalThis.process?.argv[1]&&path.resolve(globalThis.process.argv[1])===fileURLToPath(import.meta.url)){
 const r=await verificarFase2();console.log(JSON.stringify(r,null,2));if(r.fail)globalThis.process.exitCode=1;
}

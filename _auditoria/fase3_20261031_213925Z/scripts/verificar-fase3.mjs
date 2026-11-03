import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import assert from 'node:assert/strict';
import {fileURLToPath} from 'node:url';
import {parseProcedures,runAudit} from './auditar-vba.mjs';
const dir=path.resolve(path.dirname(fileURLToPath(import.meta.url)),'..');
const hash=b=>crypto.createHash('sha256').update(b).digest('hex');
const json=async p=>JSON.parse((await fs.readFile(p,'utf8')).replace(/^\uFEFF/,''));
const decode=b=>{try{return new TextDecoder('utf-8',{fatal:true}).decode(b);}catch{return new TextDecoder('windows-1252').decode(b);}};
const sig=p=>p.body.slice(0,p.body.findIndex(l=>!l.trimEnd().endsWith('_'))+1).join('\n').trim();
export async function verificarFase3(){
 const m=await json(path.join(dir,'manifesto-inicial-fase3.json'));
 const p1=await json(path.join(m.previousPhase1,'manifesto-original.json'));
 const p2=await json(path.join(m.previousPhase2,'manifesto-inicial-fase2.json'));
 const results=[];
 const check=async(name,fn)=>{try{results.push({name,status:'PASS',details:await fn()});}catch(e){results.push({name,status:'FAIL',error:e.message});}};
 await check('Originais e backups da fase 1 por SHA-256',async()=>{for(const f of p1.files){assert.equal(hash(await fs.readFile(path.join(m.root,f.relative))),f.sha256);assert.equal(hash(await fs.readFile(path.join(p1.backup,f.relative))),f.sha256);}return p1.files.length;});
 await check('Execucoes anteriores inteiras preservadas',async()=>{for(const f of m.phase2Inventory)assert.equal(hash(await fs.readFile(path.join(m.previousPhase2,f.relative))),f.sha256);for(const f of p2.previousInventory)assert.equal(hash(await fs.readFile(path.join(m.previousPhase1,f.relative))),f.sha256);return {phase1:p2.previousInventory.length,phase2:m.phase2Inventory.length};});
 await check('Backup da fase 2 confirmado',async()=>{for(const f of p2.backups)assert.equal(hash(await fs.readFile(path.join(p2.backup2,f.kind,f.relative))),f.sha256);return p2.backups.length;});
 await check('69 novos backups confirmados',async()=>{for(const f of m.newBackupRecords)assert.equal(hash(await fs.readFile(path.join(m.backup3,f.relative))),f.sha256);assert.equal(m.newBackupRecords.length,69);return 69;});
 const before=new Map(),after=new Map();
 for(const f of m.initialDevelopment.filter(f=>/\.(bas|cls|frm)$/i.test(f.relative))){before.set(f.relative,decode(await fs.readFile(path.join(m.backup3,'desenvolvimento',f.relative))));after.set(f.relative,decode(await fs.readFile(path.join(m.dev3,f.relative))));}
 await check('Somente clsOperacaoPlanilha alterada entre os fontes do sistema',async()=>{const changed=[];for(const [f,t]of before)if(after.get(f)!==t)changed.push(f);assert.deepEqual(changed,['MDL/clsOperacaoPlanilha.cls']);return changed;});
 await check('665 contratos exportados e atributos preservados',async()=>{let n=0;for(const[f,t]of before){const now=parseProcedures(after.get(f)),old=parseProcedures(t);assert.equal(now.length,old.length,f);for(const p of old){const q=now.find(q=>q.name===p.name&&q.kind===p.kind);assert.ok(q,f+':'+p.name);assert.equal(sig(q),sig(p));n++;}assert.equal(after.get(f).match(/Attribute VB_Name = "[^"]+"/)[0],t.match(/Attribute VB_Name = "[^"]+"/)[0]);}assert.equal(n,665);return n;});
 await check('Pares FRM/FRX, XML, formulas fiscais e XLSM identicos a fase 2',async()=>{for(const f of m.initialDevelopment.filter(f=>/\.(frm|frx|xlsm)$/i.test(f.relative)||f.relative.endsWith('modValidacaoXML.bas')))assert.equal(hash(await fs.readFile(path.join(m.dev3,f.relative))),f.sha256);return 'Nenhuma alteracao em controle, evento, fiscalidade ou politica de Delete.';});
 await check('Sete Delete e consumidores de recuperacao preservados',async()=>{let n=0;for(const[f,t]of before){const deletes=x=>x.split(/\r?\n/).filter(l=>/^\s*(?:Selection|c|cAnual)\..*\bDelete\b/i.test(l));assert.deepEqual(deletes(after.get(f)),deletes(t));n+=deletes(t).length;if(/Cad_NFE(?:_Saida)?\.frm$/.test(f))for(const p of parseProcedures(after.get(f)).filter(p=>/^CommandButton_Gravar(?:Filial)?_Click$/.test(p.name))){const s=p.body.join('\n');assert.ok(s.indexOf('operacao.Recuperar(diagnostico)')<s.lastIndexOf('operacao.Finalizar'));}}assert.equal(n,7);return n;});
 await check('Ciclo de vida e metadado da chave presentes',async()=>{const t=after.get('MDL/clsOperacaoPlanilha.cls');assert.ok(t.includes('Private mIniciada As Boolean'));assert.ok(t.includes('If mIniciada Or mAtiva Or mItens.Count > 0 Then'));assert.ok(t.includes('mIniciada = True'));assert.ok(t.includes('item.Add CBool(celula.Errors(xlNumberAsText).Ignore)'));assert.ok(t.includes('celula.Errors(xlNumberAsText).Ignore = CBool(item(6))'));const body=parseProcedures(t).find(p=>p.name==='Recuperar').body.join('\n');assert.ok(body.indexOf('If Not mAtiva Then')<body.indexOf('For i = mItens.Count'));return 'Contrato estatico, nao execucao VBA.';});
 await check('Suite reaproveitada tem regressao de ciclo de vida e guardas de fixture',async()=>{const t=decode(await fs.readFile(path.join(dir,'scripts/modTestesFase2.bas')));assert.ok(t.includes('AuditoriaFase3_CicloVida'));assert.ok(t.includes('InicioRepetidoRecusadoF3'));assert.ok(t.includes('AUDITORIA_VBA_FICTICIO'));return 'PREPARADO, NAO IMPORTADO, NAO EXECUTADO.';});
 await check('Quebras de linha e bytes dos 33 fontes nao alterados preservados',async()=>{for(const[f,t]of before){const raw=await fs.readFile(path.join(m.dev3,f));if(f!=='MDL/clsOperacaoPlanilha.cls')assert.equal(hash(raw),hash(await fs.readFile(path.join(m.backup3,'desenvolvimento',f))));else assert.ok(!/(?<!\r)\n/.test(raw.toString('utf8')));}return 34;});
 const audit=await runAudit(m.dev3,path.join(dir,'evidencias/estatica-final'));
 await check('Inventario sem duplicidade publica e sem nova assinatura',async()=>{assert.equal(audit.files.length,34);assert.equal(audit.procedureCount,665);assert.equal(audit.publicDuplicates.length,0);return {files:34,procedures:665,publicDuplicates:0};});
 const out={timestamp:new Date().toISOString(),modality:'Analise estatica; nao e compilacao nem execucao VBA',results,pass:results.filter(x=>x.status==='PASS').length,fail:results.filter(x=>x.status==='FAIL').length};
 await fs.writeFile(path.join(dir,'verificacao-estatica-fase3.json'),JSON.stringify(out,null,2));return out;
}
if(globalThis.process?.argv[1]&&path.resolve(globalThis.process.argv[1])===fileURLToPath(import.meta.url)){const r=await verificarFase3();console.log(JSON.stringify(r,null,2));if(r.fail)globalThis.process.exitCode=1;}

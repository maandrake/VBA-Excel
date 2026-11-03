import fs from 'node:fs/promises';
import path from 'node:path';
import crypto from 'node:crypto';
import { fileURLToPath } from 'node:url';
const here = path.dirname(fileURLToPath(import.meta.url));
const run = path.resolve(here, '..');
const root = path.resolve(run, '..', '..');
const digest = b => crypto.createHash('sha256').update(b).digest('hex');
const manifestBytes = await fs.readFile(path.join(run, 'manifesto-final-fase5.json'));
const sidecar = (await fs.readFile(path.join(run, 'manifesto-final-fase5.sha256'), 'utf8')).trim().split(/\s+/)[0];
if (digest(manifestBytes) !== sidecar) throw new Error('Manifesto SHA-256 divergente');
const manifest = JSON.parse(manifestBytes.toString('utf8'));
let files = 0;
for (const item of manifest.currentFiles) {
  const target = path.resolve(root, item.relative);
  const relative = path.relative(root, target);
  if (relative.startsWith('..') || path.isAbsolute(relative)) throw new Error('Caminho fora do workspace');
  if (digest(await fs.readFile(target)) !== item.sha256) throw new Error('Hash divergente: ' + item.relative);
  files++;
}
for (const item of manifest.originals) {
  if (digest(await fs.readFile(path.resolve(root, item.relative))) !== item.expected) throw new Error('Original alterado: ' + item.relative);
}
const scenarios = JSON.parse(await fs.readFile(path.join(run, 'evidencias', 'matriz-cenarios-executados.json'), 'utf8'));
let exact = 0;
for (const scenario of scenarios.filter(s => s.differences === 0)) {
  const folder = path.resolve(root, scenario.evidence);
  const before = await fs.readFile(path.join(folder, scenario.id + '_antes.tsv'));
  const after = await fs.readFile(path.join(folder, scenario.id + '_depois.tsv'));
  if (!before.equals(after)) throw new Error('Snapshot divergiu: ' + scenario.id);
  exact++;
}
console.log(JSON.stringify({ modality: 'Verificacao estatica de hashes e evidencias; nao executa nem compila VBA', status: 'PASS', files, originals: manifest.originals.length, exactSnapshotPairs: exact, businessScenarios: scenarios.length, counts: scenarios.reduce((a,s)=>(a[s.classification]=(a[s.classification]||0)+1,a),{}) }, null, 2));

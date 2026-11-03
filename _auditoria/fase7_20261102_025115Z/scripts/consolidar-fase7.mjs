// Leitura local das evidencias VBA; nao executa/simula macros.
import fs from 'node:fs/promises';
import path from 'node:path';
export async function readResults(dir) {
  const attempts = [];
  const latest = new Map();
  for (const file of ['leitores_resultados.tsv', 'formulas_reteste.tsv', 'frio_resultados.tsv']) {
    const lines = (await fs.readFile(path.join(dir,file),'utf16le')).replace(/^\uFEFF/,'').trimEnd().split(/\r?\n/);
    const headers = lines.shift().split('\t');
    for (const line of lines) {
      const values = line.split('\t');
      if (values.length !== headers.length) throw new Error('TSV incompleto: '+file);
      const result = Object.fromEntries(headers.map((h,i)=>[h,values[i]]));
      result.file = file;
      if (!['PASS','FAIL','BLOQUEADO','NAO EXECUTADO'].includes(result.Resultado)) throw new Error('Estado desconhecido');
      attempts.push(result); latest.set(result.Cenario,result);
    }
  }
  return {attempts,latest:[...latest.values()]};
}

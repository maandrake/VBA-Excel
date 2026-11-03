# Comparativo de armazenamento — Fase 9

Portão experimental. Nenhum motor foi escolhido ou implantado definitivamente no produto. A recomendação de **candidato de laboratório é ACE 16 via ADODB/ADOX de 64 bits já instalado**; aprovação de distribuição/licença/ambiente e esquema continua pendente.

## Ambiente efetivamente verificado

Office Click-to-Run O365BusinessRetail, x64, versão registrada 16.0.20527.20112; Excel executado 16.0/build20527. ACEOLEDB.DLL 16.0.20527.20016 em Program Files/Microsoft Office/root/VFS/ProgramFilesCommonX64/Microsoft Shared/Office16. ProgID ACE 12 e 16 abrem ACCDB em processo x64; ADO informa versão 10.0. DBMS Version=04.00.0000 é propriedade exposta pelo provedor, não identificação de uma instalação Jet 4. A DLL real pertence ao Office atual; “16.0” não prova que seja o redistribuível Access 2016.

ADOX criou banco, ADO executou BeginTrans/RollbackTrans, e o VBA repetiu operações reais no mesmo provedor. Jet 4 registrado falhou na abertura em x64: “Provedor não encontrado”. Driver ODBC 17 SQL Server está presente, mas nenhuma instância MSSQL foi encontrada na enumeração de serviços; nenhuma conexão externa foi tentada. winsqlite3.dll 3.53.4 existe; presença da DLL não prova interface ADO/ODBC para VBA, nem integração testada. Nenhum provedor SQLite foi instalado ou motor alternativo certificado. [Ambiente](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/ambiente9.json), [sondagem](./_auditoria/fase9_20261102_142753Z/evidencias/ambiente-inicial.json).

## Comparação e nível da evidência

| Critério | A — XLSM autoritativo | B — arquivos/diário externo | C — ACE local + Excel projeção |
|---|---|---|---|
| Atomicidade interna | Sem transação multicílula demonstrada; compensação F6 incompleta | Exige protocolo de geração; append de texto não cria transação entre arquivos | BeginTrans/CommitTrans/RollbackTrans reais para fatos, episódio, projeções, audit e recibo no mesmo ACCDB |
| Durabilidade | Save cancelado perde mudança somente em memória; XLSM antigo pode divergir | Parser frio F7/F8 contém órfãos/truncamentos; recuperação completa não demonstrada | Sessão Excel nova recuperou fatos/recibos/pendência; não prova energia/crash/disco cheio |
| Integridade referencial | Validação em aplicação; linha/placa não são FK | Validação de parser e relações a implementar | PK/FK/unique reais; FK composta EP/Movement testada no candidato migracao.accdb; núcleo/bench usam esquema 1 mais limitado |
| Concorrência | Cópias/OneDrive não constituem lock de domínio | Exclusão e fencing precisam implementação durável própria | Lease/CAS + bloqueio real de página/registro entre processos; órfão exige conciliação; não é controle contra acesso SQL fora do protocolo |
| Versões | Revisão no arquivo pode estar desatualizada | Manifesto e linhagem coordenados | VersionNo por episódio, RevisionNo por base, Owner/Fence por sessão; revisão da projeção recusada quando divergente |
| Recuperação | Compensação limitada e conciliação após Save/Delete | Exige conjunto coerente de dados/diário/recibos | Rollback verificado antes de commit; após commit reconstruir Excel; intenção pendente não é apagada automaticamente |
| Desempenho | F8 lista de 10 mil: 9,7135 s; verificação célula a célula | Sem benchmark novo de motor de arquivos | F9 lote10mil: 0,5041s;200 consultas verificadas: 1,7162 s. Metodologias diferentes; índice candidato não equivale a autorização |
| Backup/restauração | XLSM fechado/hash/reabertura | Copiar geração completa e testar parser/vínculos | Banco fechado copiado e reaberto em novo processo; backup on-line, compactação e desastre NÃO EXECUTADOS |
| Dependências/portabilidade | Excel/MSForms e referências existentes | APIs locais/encoding/protocolo e permissões | Office/ACE x64/ADO/ADOX comprovados nesta máquina; não presumir disponibilidade em outra nem copiar DLL |
| Segurança/privacidade | Proteção de aba não autentica usuário | ACL de todos os arquivos e backups; logs mínimos | ACL do ACCDB/pasta; sem autenticação por papel ou criptografia implantadas no ensaio. Acesso direto pode contornar a aplicação |
| Manutenção/custo operacional | Distribuição simples, alto custo de conciliação atual | Alto custo de manter durabilidade/locks/parser | Menor dependência da linha; exige esquema, testes, recuperação, licenças e atualização do Office. Não estimado custo monetário |
| Implantação | Nenhuma troca aprovada | Nenhum diário operacional aprovado | Candidato local disponível sem instalar; promoção e migração continuam P2-05/P2-06 pendentes |

A/B foram comparados com evidências existentes F6–F8 e Save real F9; não foi construído outro motor completo de arquivos apenas para repetir testes. C teve nova execução VBA e COM independente. Não há atomicidade conjunta ACCDB/XLSM.

## Suporte e licenciamento

A Microsoft informa que o redistribuível Access Database Engine 2016 terminou suporte em 14/10/2025 e aponta Microsoft 365 Access Runtime como sucessor: [Download oficial](https://www.microsoft.com/en-gb/download/details.aspx?id=54920), [Runtime/EULA](https://support.microsoft.com/en-us/access/download-and-install-microsoft-365-access-runtime). Não instalar o redistribuível legado como atalho. A titularidade/licença e o direito de implantação nas máquinas de destino precisam confirmação do responsável; a chave ProductReleaseIds não comprova licenciamento ou ativação. Não houve instalação, registro de DLL nem aceitação de EULA.

[ADO transações](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/begintrans-committrans-and-rollbacktrans-methods-ado) aplica-se à conexão/provedor. [DAO CommitTrans](https://learn.microsoft.com/en-us/office/client-developer/access/desktop-database-reference/workspace-committrans-method-dao) oferece opções próprias; esta suíte usou ADO e não demonstrou dbForceOSFlush. Propriedades observadas: User Commit Sync=True, Asynchable Commit=False, Transaction Commit Mode=0; nenhuma foi alterada. Isso não substitui teste de falha de energia. [Propriedades reais](./_auditoria/fase9_20261102_142753Z/evidencias/execucao/propriedades-ace.json).

Aceite futuro: aprovar motor/versão/interface/licenciamento e escritor único; reforçar constraints, imutabilidade e recibos; provar snapshot coerente sob escrita concorrente, recuperação de lease, restauração, permissões e falhas de commit. Até lá, ACE é recomendação técnica condicional, não banco autoritativo do produto.

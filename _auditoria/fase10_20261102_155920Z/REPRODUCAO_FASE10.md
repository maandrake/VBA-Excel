# Reprodução e restauração do laboratório F10

Não usar dados operacionais. Não abrir a entrega arquivada no Office dentro do OneDrive. O mestre efetivamente executado foi:

C:\Users\MARCOS~1\AppData\Local\Temp\vscode-agent-codex-sandbox-JTJOQt\VBA_Fase10_20261102_155920Z\Fixture_Integridade_Fase10_Mestre.xlsm

## Retomada da sessão fria bloqueada

1. Abrir somente esse mestre sintético numa instância isolada. Se o Office oferecer Habilitar Conteúdo especificamente para esse documento, essa é a permissão do arquivo de teste. Se exigir alteração de Central de Confiabilidade, política ou registro, interromper e obter autorização específica; não usar bypass.
2. Alt+F11; selecionar VBAProject do mestre; Depurar > Compilar VBAProject. Registrar literalmente erro/módulo/linha, se houver.
3. Ctrl+G; na janela Imediata executar: Call F10_Frio
4. Resultado esperado de execução: F10_COLD_DONE e frio_f10.tsv com seis linhas de teste. Copiar o log e informar mensagens originais. A classificação atual permanece BLOQUEADO enquanto não houver execução efetiva.

Na tentativa desta fase, o VBE informou macros desativadas e Arquivo > Informações informou conteúdo ativo bloqueado. Não houve alteração de segurança; a instância fria foi encerrada normalmente.

## Reprodução das suítes que criam arquivos

Não executar novamente uma suíte sobre os mesmos nomes já existentes. As rotinas usam criação sem sobrescrita. Para novo ensaio, criar pasta TEMP identificada com o marcador VBA_Fase10_20261102_155920Z, copiar/renomear a entrega para Fixture_Integridade_Fase10_Mestre.xlsm e fornecer somente os insumos sintéticos necessários. Preservar antes o diretório inteiro e hashes. F10_Schema requer legado_f9.accdb no esquema1 conhecido; F10_Core cria novo ciclo. Suítes posteriores usam backup_ciclo_f10.accdb fechado. As massas e referências estão no arquivo execucao_temp; não utilizar suas revisões intermediárias como entrega definitiva.

Ordem histórica: F10_Schema; F10_Core; F10_Transacoes; F10_AuditoriaAdversarial; F10_Projecoes; F10_BackupRestauracao; concorrência com handshake; F10_IntencaoIncompleta; F10_FKTransacao; benchmarks. As repetições após correções foram preservadas com prefixos R1_, R2_, ANTES_PREPARED_, ANTES_CUSTODIA_ e PRE_FINAL_GUARD_. Não somá-las como testes únicos novos.

Módulos finais experimentais: modPersistenciaFase10 (arquivo modLaboratorioFase10.bas), modCenariosFase10, modFronteirasFase10. Importar por UI somente se estiverem ausentes; o mestre final já os contém. Não importar duplicatas. Não usar VBProject COM bloqueado nem modificar AccessVBOM. Os demais63 componentes herdados não devem ser substituídos.

Os comandos de worker foram executados inline no PowerShell autorizado; worker_boundary_comando_executado.txt conserva o comando usado no cenário snapshot. O cenário apply só muda stage10. A política local de scripts não foi alterada. O handshake usa marker e JSON no TEMP e precisa ser preparado antes do VBA. Não reutilizar markers de outra etapa; os originais foram preservados com prefixos prepare_, snapshot_ e apply_.

## Restauração

Preservar arquivo suspeito e hash. Trabalhar em nova pasta. Conferir manifesto e baseline. Backup fechado pode reconstruir a projeção de sua revisão. Se existir revisão posterior em outra base, preservar ambas e conciliar; não sobrepor fatos novos com versão antiga. Lease órfão/RECONCILE não são eliminados por prazo. Não promover este laboratório para produção.

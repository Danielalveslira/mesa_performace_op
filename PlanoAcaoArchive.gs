/**
 * Plano_Historico e Alertas_Historico são append-only — cada criação,
 * edição, atualização de andamento, evidência e envio de alerta grava uma
 * linha nova e nenhuma delas é excluída pelo soft-delete normal (que só
 * marca "ativo=false" em Planos_Acao/Plano_Evidencias). Em qualquer volume
 * relevante de uso, essas duas abas crescem mais rápido que qualquer outra
 * e são o primeiro ponto do sistema a esbarrar nos limites práticos de
 * leitura/escrita do Google Sheets.
 *
 * archivePlanoHistory_ move os registros mais antigos que maxAgeDays para
 * uma segunda planilha ("Mesa Operacional - Arquivo Histórico", criada
 * automaticamente no Drive na primeira execução) em vez de apagá-los —
 * mantém os dados acessíveis (abra a planilha de arquivo normalmente) sem
 * inflar a planilha operacional. Não exige nenhuma infraestrutura nova
 * (BigQuery, Firestore, etc.), só outra Google Sheet.
 *
 * Não é chamada automaticamente por nenhum gatilho — é uma função
 * administrativa para rodar manualmente pelo editor do Apps Script
 * (Executar > archivePlanoHistory_) quando quiser reduzir o tamanho das
 * abas de histórico. Depois de validar o resultado algumas vezes, você
 * pode criar seu próprio gatilho de tempo (ex.: mensal) apontando para
 * ela, do mesmo jeito que setupPlanoAlertas cria o gatilho de alertas.
 */
function archivePlanoHistory_(maxAgeDays) {
  requireDeploymentOwner_();
  const ageDays = Number(maxAgeDays) > 0 ? Number(maxAgeDays) : 365;
  const cutoff = new Date(Date.now() - ageDays * 86400000);

  // Cada aba é arquivada sob o mesmo lock nomeado usado pelas operações que
  // normalmente escrevem nela (ver acquirePlanoNamedLock_ em
  // PlanoAcaoConfig.gs), para não excluir/renumerar linhas enquanto uma
  // criação de plano ou um envio de alertas está em andamento.
  const historyResult = archivePlanoSheetOlderThan_(
    PLANO_ACAO_CONFIG.SHEETS.HISTORY,
    'criado_em',
    cutoff,
    'plano-crud'
  );
  const alertResult = archivePlanoSheetOlderThan_(
    PLANO_ACAO_CONFIG.SHEETS.ALERT_HISTORY,
    'enviado_em',
    cutoff,
    'plano-alertas'
  );

  return {
    success: true,
    cutoff: Utilities.formatDate(cutoff, getPlanoTimeZone_(), 'yyyy-MM-dd'),
    archiveSpreadsheetUrl: getPlanoArchiveSpreadsheet_().getUrl(),
    archived: [historyResult, alertResult],
  };
}

function archivePlanoSheetOlderThan_(sheetName, dateField, cutoff, lockDomain) {
  const lock = requirePlanoNamedLock_(lockDomain, 30000);
  try {
    const records = readPlanoRecords_(sheetName).filter((record) => {
      const value = record[dateField];
      const date = value instanceof Date ? value : new Date(value);
      return !Number.isNaN(date.getTime()) && date < cutoff;
    });
    if (!records.length) return { sheet: sheetName, archived: 0 };

    const archiveSheet = ensurePlanoSheet_(
      getPlanoArchiveSpreadsheet_(),
      sheetName,
      PLANO_ACAO_CONFIG.HEADERS[sheetName]
    );
    appendPlanoRecords_(archiveSheet, records);
    const removed = deletePlanoRows_(getPlanoSheet_(sheetName), records.map((record) => record.__row));
    SpreadsheetApp.flush();
    return { sheet: sheetName, archived: removed };
  } finally {
    lock.release();
  }
}

function getPlanoArchiveSpreadsheet_() {
  const properties = PropertiesService.getScriptProperties();
  const id = properties.getProperty('PLANO_ACAO_ARCHIVE_SPREADSHEET_ID');
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (error) {
      // A planilha de arquivo foi excluída ou o acesso foi perdido; recria abaixo.
    }
  }

  const archive = SpreadsheetApp.create('Mesa Operacional - Arquivo Histórico');
  properties.setProperty('PLANO_ACAO_ARCHIVE_SPREADSHEET_ID', archive.getId());
  return archive;
}

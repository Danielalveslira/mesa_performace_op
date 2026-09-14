/**
 * Erro destinado ao usuário final. Sua mensagem é sempre segura para exibição
 * no navegador. Qualquer outro erro (bug, falha de API do Google, etc.) é
 * tratado por planoRunPublic_ como técnico e nunca chega ao navegador como está.
 */
class PlanoUserError extends Error {
  constructor(message) {
    super(message);
    this.name = 'PlanoUserError';
  }
}

/**
 * Executa uma função de entrada pública (chamada via google.script.run) com uma
 * borda de erro padronizada: mensagens de negócio (PlanoUserError) chegam ao
 * navegador como estão; qualquer outro erro é registrado por completo no
 * Stackdriver (Apps Script > Execuções) e substituído por uma mensagem genérica,
 * evitando vazar detalhes técnicos (nomes de aba, stack trace, etc.) ao usuário.
 */
function planoRunPublic_(label, fn) {
  try {
    return fn();
  } catch (error) {
    if (error instanceof PlanoUserError) throw error;
    console.error(`Falha inesperada em ${label}: ${(error && error.stack) || error}`);
    throw new Error('Ocorreu um erro inesperado. Tente novamente em instantes.');
  }
}

/**
 * LockService só oferece um mutex por escopo fixo (script/usuário/documento):
 * não existe um "getScriptLock('minha-chave')". Antes desta mudança, TODA
 * mutação do sistema — criação/edição de plano, upload de evidência e
 * envio de alertas — disputava o mesmo LockService.getScriptLock() global,
 * então processar alertas bloqueava (e era bloqueado por) qualquer edição
 * de plano em uma regional completamente diferente.
 *
 * Isto implementa um mutex nomeado sobre o próprio script lock: o script
 * lock é seguro (fornece exclusão mútua atômica de verdade) e é mantido
 * apenas pelo tempo mínimo de checar/gravar uma chave no CacheService — não
 * pelo tempo da operação inteira. O trabalho real acontece com o script
 * lock já liberado, então operações em domínios diferentes (ex.: 'plano-
 * crud' vs. 'plano-alertas') deixam de se bloquear mutuamente.
 *
 * Dentro de um mesmo domínio o comportamento de serialização é idêntico ao
 * lock global anterior — nenhuma das travas de segurança já existentes
 * (renumeração de linhas em exclusões, __row capturado antes de escrever)
 * foi alterada, porque createPlanoAcao_impl_/updatePlanoAcao_impl_/
 * addPlanoAcaoUpdate_impl_/uploadPlanoEvidence_impl_/deletePlanoEvidence_impl_/
 * deletePlanoAcao_impl_ continuam todos no mesmo domínio 'plano-crud'.
 *
 * O CacheService tem TTL próprio (60s) como rede de segurança: se uma
 * execução travar/expirar sem liberar (timeout do Apps Script, erro não
 * tratado antes do finally), o lock nomeado se autolibera em até 60s em vez
 * de ficar preso para sempre.
 */
function acquirePlanoNamedLock_(domain, waitMs) {
  const cache = CacheService.getScriptCache();
  const key = `plano-named-lock:${domain}`;
  const deadline = Date.now() + Math.max(0, Number(waitMs) || 0);

  for (;;) {
    const scriptLock = LockService.getScriptLock();
    scriptLock.waitLock(5000);
    try {
      if (!cache.get(key)) {
        cache.put(key, '1', 60);
        return { release: () => cache.remove(key) };
      }
    } finally {
      scriptLock.releaseLock();
    }
    if (Date.now() >= deadline) return null;
    Utilities.sleep(300);
  }
}

/**
 * Variante de acquirePlanoNamedLock_ que lança PlanoUserError em vez de
 * devolver null, para os pontos de entrada que sempre precisam do lock
 * para prosseguir (equivalente ao antigo lock.waitLock(30000) + uso direto).
 */
function requirePlanoNamedLock_(domain, waitMs) {
  const lock = acquirePlanoNamedLock_(domain, waitMs);
  if (!lock) {
    throw new PlanoUserError('O sistema está ocupado processando outra operação. Tente novamente em instantes.');
  }
  return lock;
}

const PLANO_ACAO_CONFIG = Object.freeze({
  SCHEMA_VERSION: '2',
  SCHEMA_VERSION_PROPERTY: 'PLANO_ACAO_SCHEMA_VERSION',
  EVIDENCE_FOLDER_PROPERTY: 'PLANO_ACAO_EVIDENCE_FOLDER_ID',
  MAX_EVIDENCE_BYTES: 4 * 1024 * 1024,
  MAX_EVIDENCE_COUNT_PER_PLAN: 20,
  MAX_EVIDENCE_TOTAL_BYTES_PER_PLAN: 40 * 1024 * 1024,
  MAX_LINKS: 20,
  MAX_LINK_LENGTH: 2048,
  TEXT_LIMITS: Object.freeze({
    product: 80,
    action: 2000,
    execution: 5000,
    responsibleName: 150,
    email: 254,
    pendingReason: 2000,
    nextStep: 2000,
    updateSummary: 2000,
    historyReason: 1000,
  }),
  ALLOWED_EVIDENCE_TYPES: Object.freeze(['image/jpeg', 'image/png', 'image/webp']),
  PUBLIC_FIELDS: Object.freeze({
    PLAN: Object.freeze([
      'id', 'cidade', 'regional', 'produto', 'periodo_analisado', 'periodo_comparacao',
      'base_anterior', 'base_atual', 'diferenca', 'variacao_pct', 'prioridade', 'o_que',
      'como', 'responsavel', 'responsavel_email', 'prazo', 'status', 'percentual_conclusao',
      'pendencia_motivo', 'proximo_passo', 'links_evidencias', 'criado_por', 'criado_em',
      'atualizado_por', 'atualizado_em', 'versao', 'ativo',
    ]),
    // Usado na listagem (getPlanoAcaoBootstrap/getAppBootstrap), que hoje
    // carrega todos os planos acessíveis de uma vez: omite os campos de
    // texto longo que só o formulário de edição e a tela de detalhe usam
    // (como, pendencia_motivo, proximo_passo, links_evidencias — até
    // 5000+2000+2000+20*2048 caracteres por plano) e que nunca são lidos a
    // partir de um item da lista no cliente, apenas a partir de
    // getPlanoAcaoDetail. Conferido em PlanoAcaoClient.html: o único ponto
    // que pré-preenche o formulário de edição com esses campos
    // (planoOpenForm) sempre recebe detail.plan, nunca um item de
    // planoState.plans. Mantém todo o resto (inclusive campos hoje sem uso
    // na UI, como base_anterior/diferenca/variacao_pct) para não quebrar
    // nada que dependa deles no futuro sem motivo — só corta o que
    // comprovadamente pesa mais e não é lido na lista.
    PLAN_LIST: Object.freeze([
      'id', 'cidade', 'regional', 'periodo_analisado', 'periodo_comparacao',
      'base_atual', 'prioridade', 'o_que', 'responsavel', 'prazo', 'status',
      'percentual_conclusao', 'criado_em', 'atualizado_em', 'versao', 'ativo',
    ]),
    UPDATE: Object.freeze([
      'id', 'tipo', 'resumo', 'status', 'percentual_conclusao', 'pendencia_motivo',
      'proximo_passo', 'novo_prazo', 'links', 'autor', 'criado_em', 'versao_plano_resultante',
    ]),
    HISTORY: Object.freeze(['id', 'operacao', 'motivo', 'autor', 'criado_em']),
    EVIDENCE: Object.freeze(['id', 'nome', 'tipo', 'tamanho', 'autor', 'criado_em']),
  }),
  SHEETS: Object.freeze({
    PLANS: 'Planos_Acao',
    UPDATES: 'Plano_Atualizacoes',
    HISTORY: 'Plano_Historico',
    EVIDENCE: 'Plano_Evidencias',
    USERS: 'Usuarios_Permissoes',
    ALERT_CONFIG: 'Alertas_Config',
    ALERT_HISTORY: 'Alertas_Historico',
  }),
  STATUS: Object.freeze(['nao_iniciado', 'em_andamento', 'pendente', 'concluido', 'cancelado']),
  PRIORITIES: Object.freeze(['baixa', 'media', 'alta', 'critica']),
  ROLES: Object.freeze(['visualizador', 'editor', 'administrador']),
  HEADERS: Object.freeze({
    Planos_Acao: Object.freeze([
      'id', 'operacao_id', 'cidade', 'regional', 'produto', 'periodo_analisado', 'periodo_comparacao',
      'base_anterior', 'base_atual', 'diferenca', 'variacao_pct', 'prioridade', 'o_que',
      'como', 'responsavel', 'responsavel_email', 'prazo', 'status', 'percentual_conclusao', 'pendencia_motivo',
      'proximo_passo', 'links_evidencias_json', 'contexto_criacao_json', 'criado_por',
      'criado_em', 'atualizado_por', 'atualizado_em', 'versao', 'ativo',
    ]),
    Plano_Atualizacoes: Object.freeze([
      'id', 'operacao_id', 'plano_id', 'tipo', 'resumo', 'status', 'percentual_conclusao',
      'pendencia_motivo', 'proximo_passo', 'novo_prazo', 'links_json', 'autor',
      'criado_em', 'versao_plano_resultante',
    ]),
    Plano_Historico: Object.freeze([
      'id', 'operacao_id', 'plano_id', 'operacao', 'antes_json', 'depois_json', 'motivo', 'autor', 'criado_em',
    ]),
    // 'url' foi removida: nunca foi preenchida nem lida (o conteúdo é servido
    // via getPlanoEvidenceContent a partir do drive_file_id). Planilhas já
    // existentes mantêm a coluna antiga sem uso; novas planilhas não a criam.
    Plano_Evidencias: Object.freeze([
      'id', 'plano_id', 'atualizacao_id', 'drive_file_id', 'nome', 'tipo', 'tamanho',
      'checksum_sha256', 'autor', 'criado_em', 'ativo',
    ]),
    Usuarios_Permissoes: Object.freeze([
      'email', 'nome', 'papel', 'regionais_json', 'permissoes_json', 'recebe_alertas_gestao', 'ativo',
      'criado_em', 'atualizado_em',
    ]),
    Alertas_Config: Object.freeze([
      'chave', 'valor', 'descricao', 'atualizado_em',
    ]),
    Alertas_Historico: Object.freeze([
      'id', 'plano_id', 'tipo', 'referencia', 'chave_unica', 'destinatario',
      'assunto', 'enviado_em', 'status', 'erro',
    ]),
  }),
});

function setupPlanoAcao_() {
  requireDeploymentOwner_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const spreadsheet = getSpreadsheet_();
    Object.keys(PLANO_ACAO_CONFIG.HEADERS).forEach((sheetName) => {
      ensurePlanoSheet_(spreadsheet, sheetName, PLANO_ACAO_CONFIG.HEADERS[sheetName]);
    });

    bootstrapPlanoAdmin_(spreadsheet);
    SpreadsheetApp.flush();
    PropertiesService.getScriptProperties().setProperty(
      PLANO_ACAO_CONFIG.SCHEMA_VERSION_PROPERTY,
      PLANO_ACAO_CONFIG.SCHEMA_VERSION
    );

    return {
      success: true,
      message: 'Módulo de planos de ação configurado com sucesso.',
      sheets: Object.keys(PLANO_ACAO_CONFIG.HEADERS),
    };
  } finally {
    lock.releaseLock();
  }
}

function ensurePlanoRuntimeSchema_() {
  const properties = PropertiesService.getScriptProperties();
  if (
    properties.getProperty(PLANO_ACAO_CONFIG.SCHEMA_VERSION_PROPERTY)
    === PLANO_ACAO_CONFIG.SCHEMA_VERSION
  ) {
    return;
  }

  const spreadsheet = getSpreadsheet_();
  [
    PLANO_ACAO_CONFIG.SHEETS.PLANS,
    PLANO_ACAO_CONFIG.SHEETS.UPDATES,
    PLANO_ACAO_CONFIG.SHEETS.HISTORY,
  ].forEach((sheetName) => {
    ensurePlanoSheet_(spreadsheet, sheetName, PLANO_ACAO_CONFIG.HEADERS[sheetName]);
  });
  SpreadsheetApp.flush();
  properties.setProperty(
    PLANO_ACAO_CONFIG.SCHEMA_VERSION_PROPERTY,
    PLANO_ACAO_CONFIG.SCHEMA_VERSION
  );
}

function ensurePlanoSheet_(spreadsheet, sheetName, requiredHeaders) {
  let sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) sheet = spreadsheet.insertSheet(sheetName);

  const lastColumn = sheet.getLastColumn();
  const existingHeaders = lastColumn > 0
    ? sheet.getRange(1, 1, 1, lastColumn).getValues()[0].map((value) => String(value || '').trim())
    : [];
  const existingNormalized = new Set(existingHeaders.map(normalizeHeader_));
  const missingHeaders = requiredHeaders.filter((header) => !existingNormalized.has(normalizeHeader_(header)));

  if (!existingHeaders.length) {
    sheet.getRange(1, 1, 1, requiredHeaders.length).setValues([requiredHeaders]);
  } else if (missingHeaders.length) {
    sheet.getRange(1, existingHeaders.length + 1, 1, missingHeaders.length).setValues([missingHeaders]);
  }

  formatPlanoSheet_(sheet, sheetName);
  protectPlanoSheet_(sheet);
  return sheet;
}

function protectPlanoSheet_(sheet) {
  const description = `Mesa Operacional - ${sheet.getName()}`;
  const existingProtection = sheet
    .getProtections(SpreadsheetApp.ProtectionType.SHEET)
    .find((protection) => protection.getDescription() === description);
  const protection = existingProtection || sheet.protect().setDescription(description);
  const owner = Session.getEffectiveUser();
  const ownerEmail = String(owner.getEmail() || '').trim().toLowerCase();

  protection.setWarningOnly(false);
  protection.addEditor(owner);

  const removableEditors = protection.getEditors().filter((editor) =>
    String(editor.getEmail() || '').trim().toLowerCase() !== ownerEmail
  );
  if (removableEditors.length) protection.removeEditors(removableEditors);
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
}

function formatPlanoSheet_(sheet, sheetName) {
  const lastColumn = sheet.getLastColumn();
  if (!lastColumn) return;

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, lastColumn)
    .setBackground('#0f1b32')
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setWrap(true);

  if (!sheet.getFilter() && sheet.getMaxRows() > 1) {
    sheet.getRange(1, 1, sheet.getMaxRows(), lastColumn).createFilter();
  }

  const headers = createHeaderMap_(sheet.getRange(1, 1, 1, lastColumn).getValues()[0]);
  setPlanoColumnWidth_(sheet, headers, 'id', 220);
  setPlanoColumnWidth_(sheet, headers, 'cidade', 190);
  setPlanoColumnWidth_(sheet, headers, 'o_que', 280);
  setPlanoColumnWidth_(sheet, headers, 'como', 320);
  setPlanoColumnWidth_(sheet, headers, 'resumo', 320);
  setPlanoColumnWidth_(sheet, headers, 'responsavel', 190);
  setPlanoColumnWidth_(sheet, headers, 'responsavel_email', 230);
  setPlanoColumnWidth_(sheet, headers, 'proximo_passo', 260);
  setPlanoColumnWidth_(sheet, headers, 'descricao', 360);
  setPlanoColumnWidth_(sheet, headers, 'destinatario', 230);
  setPlanoColumnWidth_(sheet, headers, 'erro', 320);

  if (sheetName === PLANO_ACAO_CONFIG.SHEETS.PLANS) {
    setPlanoNumberFormat_(sheet, headers, ['base_anterior', 'base_atual', 'diferenca'], '#,##0');
    setPlanoNumberFormat_(sheet, headers, ['variacao_pct'], '0.00%');
    setPlanoNumberFormat_(sheet, headers, ['prazo'], 'yyyy-mm-dd');
    setPlanoNumberFormat_(sheet, headers, ['criado_em', 'atualizado_em'], 'yyyy-mm-dd hh:mm:ss');
    setPlanoValidation_(sheet, headers, 'status', PLANO_ACAO_CONFIG.STATUS);
    setPlanoValidation_(sheet, headers, 'prioridade', PLANO_ACAO_CONFIG.PRIORITIES);
  } else if (sheetName === PLANO_ACAO_CONFIG.SHEETS.USERS) {
    setPlanoValidation_(sheet, headers, 'papel', PLANO_ACAO_CONFIG.ROLES);
    setPlanoValidation_(sheet, headers, 'recebe_alertas_gestao', ['TRUE', 'FALSE']);
    setPlanoValidation_(sheet, headers, 'ativo', ['TRUE', 'FALSE']);
    setPlanoNumberFormat_(sheet, headers, ['criado_em', 'atualizado_em'], 'yyyy-mm-dd hh:mm:ss');
  } else if (sheetName === PLANO_ACAO_CONFIG.SHEETS.ALERT_CONFIG) {
    setPlanoNumberFormat_(sheet, headers, ['atualizado_em'], 'yyyy-mm-dd hh:mm:ss');
  } else if (sheetName === PLANO_ACAO_CONFIG.SHEETS.ALERT_HISTORY) {
    setPlanoNumberFormat_(sheet, headers, ['enviado_em'], 'yyyy-mm-dd hh:mm:ss');
  } else {
    setPlanoNumberFormat_(sheet, headers, ['novo_prazo'], 'yyyy-mm-dd');
    setPlanoNumberFormat_(sheet, headers, ['criado_em', 'atualizado_em'], 'yyyy-mm-dd hh:mm:ss');
  }
}

function setPlanoColumnWidth_(sheet, headers, header, width) {
  const index = headers.get(normalizeHeader_(header));
  if (index !== undefined) sheet.setColumnWidth(index + 1, width);
}

function setPlanoNumberFormat_(sheet, headers, names, format) {
  names.forEach((name) => {
    const index = headers.get(normalizeHeader_(name));
    if (index !== undefined && sheet.getMaxRows() > 1) {
      sheet.getRange(2, index + 1, sheet.getMaxRows() - 1, 1).setNumberFormat(format);
    }
  });
}

function setPlanoValidation_(sheet, headers, header, values) {
  const index = headers.get(normalizeHeader_(header));
  if (index === undefined || sheet.getMaxRows() <= 1) return;

  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true)
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, index + 1, sheet.getMaxRows() - 1, 1).setDataValidation(rule);
}

function bootstrapPlanoAdmin_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(PLANO_ACAO_CONFIG.SHEETS.USERS);
  if (sheet.getLastRow() > 1) return;

  const email = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) {
    throw new PlanoUserError('Não foi possível identificar o e-mail para criar o administrador inicial.');
  }

  const now = new Date();
  appendPlanoRecord_(sheet, {
    email,
    nome: email.split('@')[0],
    papel: 'administrador',
    regionais_json: '[]',
    permissoes_json: '{}',
    recebe_alertas_gestao: true,
    ativo: true,
    criado_em: now,
    atualizado_em: now,
  });
}

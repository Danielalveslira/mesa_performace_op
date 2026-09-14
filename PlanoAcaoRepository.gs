let planoTimeZoneCache_ = '';

function getPlanoSheet_(sheetName) {
  const sheet = getSpreadsheet_().getSheetByName(sheetName);
  if (!sheet) {
    throw new PlanoUserError('O módulo de planos de ação ainda não foi configurado. Execute setupPlanoAcao_.');
  }
  return sheet;
}

function readPlanoRecords_(sheetName) {
  const sheet = getPlanoSheet_(sheetName);
  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];

  const headers = values[0].map((header) => String(header || '').trim());
  return values.slice(1)
    .map((row, index) => createPlanoRecord_(headers, row, index + 2))
    .filter(hasPlanoRecordData_);
}

function readPlanoRecordsByField_(sheetName, fieldName, fieldValue) {
  const sheet = getPlanoSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  const lastColumn = sheet.getLastColumn();
  if (lastRow < 2 || !lastColumn) return [];

  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0]
    .map((header) => String(header || '').trim());
  const fieldIndex = headers.findIndex((header) => normalizeHeader_(header) === normalizeHeader_(fieldName));
  if (fieldIndex < 0) throw new PlanoUserError(`Coluna obrigatória não encontrada: ${fieldName}`);

  const matches = sheet
    .getRange(2, fieldIndex + 1, lastRow - 1, 1)
    .createTextFinder(String(fieldValue || ''))
    .matchEntireCell(true)
    .findAll();
  if (!matches.length) return [];

  const rowNumbers = Array.from(new Set(matches.map((range) => range.getRow()))).sort((a, b) => a - b);
  const rowRanges = sheet.getRangeList(
    rowNumbers.map((rowNumber) => sheet.getRange(rowNumber, 1, 1, lastColumn).getA1Notation())
  ).getRanges();

  return rowRanges
    .map((range, index) => createPlanoRecord_(headers, range.getValues()[0], rowNumbers[index]))
    .filter(hasPlanoRecordData_);
}

/**
 * Lê Plano_Atualizacoes, Plano_Historico e Plano_Evidencias filtradas por
 * plano_id em uma única chamada HTTP (Sheets.Spreadsheets.Values.batchGet),
 * em vez de 3 idas e voltas separadas via SpreadsheetApp — cada uma delas
 * já custando 3 sub-chamadas (cabeçalho + TextFinder + getRangeList) em
 * readPlanoRecordsByField_. Usado por getPlanoAcaoDetail_impl_, o caminho
 * mais executado do módulo (abrir qualquer plano).
 *
 * Requer o serviço avançado "Sheets API" habilitado no projeto (já
 * declarado em appsscript.json > dependencies.enabledAdvancedServices; em
 * projetos antigos pode ser preciso habilitar manualmente em Serviços, na
 * barra lateral do editor do Apps Script). Se o serviço não estiver
 * disponível por qualquer motivo, cai para o caminho antigo automaticamente
 * — nenhuma funcionalidade depende de tê-lo habilitado, só a performance.
 */
function readPlanoRelatedRecordsBatch_(planId) {
  if (typeof Sheets === 'undefined' || !Sheets.Spreadsheets || !Sheets.Spreadsheets.Values) {
    return readPlanoRelatedRecordsFallback_(planId);
  }

  const sheetKeys = [
    ['updates', PLANO_ACAO_CONFIG.SHEETS.UPDATES],
    ['history', PLANO_ACAO_CONFIG.SHEETS.HISTORY],
    ['evidence', PLANO_ACAO_CONFIG.SHEETS.EVIDENCE],
  ];

  try {
    const spreadsheetId = getSpreadsheet_().getId();
    const timezone = getPlanoTimeZone_();
    const response = Sheets.Spreadsheets.Values.batchGet(spreadsheetId, {
      ranges: sheetKeys.map(([, sheetName]) => sheetName),
      // UNFORMATTED_VALUE evita risco de separador de milhar em campos
      // numéricos (tamanho, percentual_conclusao); FORMATTED_STRING faz as
      // células de data virem como texto "yyyy-MM-dd[ HH:mm:ss]" (o mesmo
      // formato que ensurePlanoSheet_ aplica a essas colunas), em vez de
      // número de série — convertido de volta para Date logo abaixo.
      valueRenderOption: 'UNFORMATTED_VALUE',
      dateTimeRenderOption: 'FORMATTED_STRING',
    });
    const valueRanges = response.valueRanges || [];
    const result = {};
    sheetKeys.forEach(([key], index) => {
      const values = (valueRanges[index] && valueRanges[index].values) || [];
      result[key] = valuesToPlanoRecords_(values)
        .filter((record) => String(record.plano_id || '') === String(planId || ''))
        .map((record) => parsePlanoBatchDateFields_(record, timezone));
    });
    return result;
  } catch (error) {
    console.warn(
      `Falha ao usar o Sheets Advanced Service no detalhe do plano, usando leitura padrão: ${
        error && error.message ? error.message : error
      }`
    );
    return readPlanoRelatedRecordsFallback_(planId);
  }
}

function readPlanoRelatedRecordsFallback_(planId) {
  return {
    updates: readPlanoRecordsByField_(PLANO_ACAO_CONFIG.SHEETS.UPDATES, 'plano_id', planId),
    history: readPlanoRecordsByField_(PLANO_ACAO_CONFIG.SHEETS.HISTORY, 'plano_id', planId),
    evidence: readPlanoRecordsByField_(PLANO_ACAO_CONFIG.SHEETS.EVIDENCE, 'plano_id', planId),
  };
}

function valuesToPlanoRecords_(values) {
  if (!values || !values.length) return [];
  const headers = values[0].map((header) => String(header || '').trim());
  return values.slice(1)
    .map((row, index) => createPlanoRecord_(headers, row, index + 2))
    .filter(hasPlanoRecordData_);
}

// Campos de data conhecidos das 3 abas lidas por readPlanoRelatedRecordsBatch_
// (Plano_Atualizacoes, Plano_Historico, Plano_Evidencias) e o formato que
// ensurePlanoSheet_/formatPlanoSheet_ garante para cada um.
const PLANO_BATCH_DATE_FIELDS_ = Object.freeze({
  criado_em: 'yyyy-MM-dd HH:mm:ss',
  novo_prazo: 'yyyy-MM-dd',
});

/**
 * Converte os campos de data (recebidos como texto formatado do batchGet)
 * de volta para objetos Date, para serializePlanoRecord_ tratá-los como já
 * trata qualquer outra leitura via SpreadsheetApp. Se o texto não bater com
 * o formato esperado — planilha antiga sem o número de formato aplicado a
 * essa linha, por exemplo — lança um erro em vez de gravar um valor
 * inconsistente; o catch em readPlanoRelatedRecordsBatch_ captura isso e
 * usa o caminho de leitura padrão para esta requisição.
 */
function parsePlanoBatchDateFields_(record, timezone) {
  Object.keys(PLANO_BATCH_DATE_FIELDS_).forEach((field) => {
    if (!Object.prototype.hasOwnProperty.call(record, field)) return;
    const value = record[field];
    if (value === '' || value === null || value === undefined) return;

    const format = PLANO_BATCH_DATE_FIELDS_[field];
    const text = String(value).trim();
    const date = Utilities.parseDate(text, timezone, format);
    if (Number.isNaN(date.getTime()) || Utilities.formatDate(date, timezone, format) !== text) {
      throw new Error(`Formato de data inesperado em ${field}: ${text}`);
    }
    record[field] = date;
  });
  return record;
}

function deletePlanoRecordsByField_(sheetName, fieldName, fieldValue) {
  const records = readPlanoRecordsByField_(sheetName, fieldName, fieldValue);
  return deletePlanoRows_(getPlanoSheet_(sheetName), records.map((record) => record.__row));
}

function deletePlanoRows_(sheet, rowNumbers) {
  const rows = Array.from(new Set(
    (rowNumbers || []).map(Number).filter((row) => Number.isInteger(row) && row >= 2)
  )).sort((a, b) => b - a);
  if (!rows.length) return 0;

  let blockStart = rows[0];
  let blockEnd = rows[0];
  const blocks = [];

  for (let index = 1; index < rows.length; index += 1) {
    const row = rows[index];
    if (row === blockStart - 1) {
      blockStart = row;
    } else {
      blocks.push({ start: blockStart, count: blockEnd - blockStart + 1 });
      blockStart = row;
      blockEnd = row;
    }
  }
  blocks.push({ start: blockStart, count: blockEnd - blockStart + 1 });

  blocks.forEach((block) => sheet.deleteRows(block.start, block.count));
  return rows.length;
}

function createPlanoRecord_(headers, row, rowNumber) {
  const record = { __row: rowNumber };
  headers.forEach((header, column) => {
    if (header) record[header] = row[column];
  });
  return record;
}

function hasPlanoRecordData_(record) {
  return Object.keys(record).some((key) => key !== '__row' && record[key] !== '');
}

function appendPlanoRecord_(sheet, record) {
  return appendPlanoRecords_(sheet, [record])[0];
}

/**
 * Grava vários registros em uma única chamada de planilha (1 leitura de
 * cabeçalho + 1 escrita em lote), em vez de uma chamada de API por registro.
 * Usado, por exemplo, no histórico de alertas, onde dezenas de linhas podem
 * ser geradas na mesma execução. Retorna os números das linhas gravadas, na
 * mesma ordem dos registros recebidos.
 */
function appendPlanoRecords_(sheet, records) {
  if (!records || !records.length) return [];
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const rows = records.map((record) =>
    headers.map((header) => Object.prototype.hasOwnProperty.call(record, header) ? record[header] : '')
  );
  const firstRow = sheet.getLastRow() + 1;
  sheet.getRange(firstRow, 1, rows.length, lastColumn).setValues(rows);
  return rows.map((row, index) => firstRow + index);
}

function updatePlanoRecord_(sheet, rowNumber, changes) {
  const lastColumn = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastColumn).getValues()[0];
  const range = sheet.getRange(rowNumber, 1, 1, lastColumn);
  const row = range.getValues()[0];
  headers.forEach((header, index) => {
    if (Object.prototype.hasOwnProperty.call(changes, header)) row[index] = changes[header];
  });
  range.setValues([row]);
}

function findPlanoRecordById_(sheetName, id) {
  const normalizedId = String(id || '').trim();
  if (!normalizedId || normalizedId.length > 100) return null;
  return readPlanoRecordsByField_(sheetName, 'id', normalizedId)[0] || null;
}

function serializePlanoRecord_(record) {
  if (!record) return null;
  const timezone = getPlanoTimeZone_();
  const serialized = {};

  Object.keys(record).forEach((key) => {
    if (key === '__row' || key === 'operacao_id') return;
    const value = record[key];
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      if (/^periodo_(analisado|comparacao)$/.test(key)) {
        serialized[key] = Utilities.formatDate(value, timezone, 'yyyy-MM');
      } else if (/(^prazo$|^novo_prazo$)/.test(key)) {
        serialized[key] = Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
      } else {
        serialized[key] = Utilities.formatDate(value, timezone, "yyyy-MM-dd'T'HH:mm:ss");
      }
    } else {
      serialized[key] = value;
    }
  });

  ['links_evidencias_json', 'contexto_criacao_json', 'links_json'].forEach((key) => {
    if (Object.prototype.hasOwnProperty.call(serialized, key)) {
      serialized[key.replace(/_json$/, '')] = parsePlanoJson_(serialized[key], key.includes('links') ? [] : {});
      delete serialized[key];
    }
  });
  if (Object.prototype.hasOwnProperty.call(serialized, 'versao')) {
    serialized.versao = Number(serialized.versao) || 1;
  }
  if (Object.prototype.hasOwnProperty.call(serialized, 'percentual_conclusao')) {
    serialized.percentual_conclusao = Number(serialized.percentual_conclusao) || 0;
  }
  if (Object.prototype.hasOwnProperty.call(serialized, 'ativo')) {
    serialized.ativo = toPlanoBoolean_(serialized.ativo, true);
  }
  return serialized;
}

function serializePlanoPublicRecord_(record, allowedFields) {
  const serialized = serializePlanoRecord_(record);
  if (!serialized) return null;
  return (allowedFields || []).reduce((result, field) => {
    if (Object.prototype.hasOwnProperty.call(serialized, field)) {
      result[field] = serialized[field];
    }
    return result;
  }, {});
}

function getPlanoTimeZone_() {
  if (!planoTimeZoneCache_) {
    planoTimeZoneCache_ = normalizeTimeZone_(getSpreadsheet_().getSpreadsheetTimeZone());
  }
  return planoTimeZoneCache_;
}

function parsePlanoJson_(value, fallback) {
  if (value && typeof value === 'object') return value;
  try {
    return JSON.parse(String(value || ''));
  } catch (error) {
    return fallback;
  }
}

function toPlanoBoolean_(value, fallback) {
  if (value === true || String(value).toLowerCase() === 'true' || Number(value) === 1) return true;
  if (value === false || String(value).toLowerCase() === 'false' || Number(value) === 0) return false;
  return fallback;
}

function addPlanoHistory_(planId, operation, before, after, reason, email, operationId) {
  const history = {
    id: Utilities.getUuid(),
    operacao_id: String(operationId || '').trim(),
    plano_id: planId,
    operacao: operation,
    antes_json: JSON.stringify(before || {}),
    depois_json: JSON.stringify(after || {}),
    motivo: String(reason || '').trim(),
    autor: email,
    criado_em: new Date(),
  };
  history.__row = appendPlanoRecord_(getPlanoSheet_(PLANO_ACAO_CONFIG.SHEETS.HISTORY), history);
  return history;
}

const CONFIG = Object.freeze({
  SOURCE_SHEET: 'vw_indicadores_cidades',
  REGION_SHEET: 'Dados',
  GOALS_SHEET: 'Base_METAS_importada',
  // v9: removida a validação de metas (goalValidation saiu do payload por
  // cidade) — versão trocada para não servir um payload em cache com o
  // formato antigo enquanto o TTL de 300s do cache anterior não expira.
  CACHE_KEY: 'dashboard-data-v9',
  CACHE_REVISION_PROPERTY: 'DASHBOARD_CACHE_REVISION',
  CACHE_SECONDS: 300,
});

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Mesa Operacional — Base Ativa')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include_(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getAppBootstrap(forceRefresh, selectedCurrentMonth, selectedPreviousMonth) {
  return planoRunPublic_('getAppBootstrap', () => {
    const user = getAppUser_();
    return {
      user: publicPlanoUser_(user),
      permissions: createPlanoPermissionFlags_(user),
      dashboard: getDashboardDataForUser_(
        user,
        forceRefresh,
        selectedCurrentMonth,
        selectedPreviousMonth
      ),
      plans: createPlanoAcaoBootstrap_(user),
    };
  });
}

function getDashboardDataForUser_(user, forceRefresh, selectedCurrentMonth, selectedPreviousMonth) {
  const dashboard = getRawDashboardData_(
    forceRefresh,
    selectedCurrentMonth,
    selectedPreviousMonth
  );
  return filterDashboardForUser_(dashboard, user);
}

function getRawDashboardData_(forceRefresh, selectedCurrentMonth, selectedPreviousMonth) {
  const cache = CacheService.getScriptCache();
  const currentMonth = normalizeMonthKey_(selectedCurrentMonth);
  const previousMonth = normalizeMonthKey_(selectedPreviousMonth);
  const cacheRevision = forceRefresh
    ? invalidateDashboardCache_()
    : getDashboardCacheRevision_();
  const cacheKey = createCacheKey_(cacheRevision, currentMonth, previousMonth);

  if (!forceRefresh) {
    const cached = cache.get(cacheKey);
    // Cidade, população e metas mudam raramente; não vale reler as abas a cada
    // acerto de cache (isso tornava até os acertos de cache tão caros quanto
    // uma reconstrução completa, e acoplava o caminho rápido a mais uma aba).
    // Essas informações já se atualizam quando o cache expira
    // (CONFIG.CACHE_SECONDS) ou quando alguém força um refresh.
    if (cached) return JSON.parse(cached);
  }

  const spreadsheet = getSpreadsheet_();
  const sourceSheet = getRequiredSheet_(spreadsheet, CONFIG.SOURCE_SHEET);
  const regionSheet = getRequiredSheet_(spreadsheet, CONFIG.REGION_SHEET);
  const goalsSheet = getRequiredSheet_(spreadsheet, CONFIG.GOALS_SHEET);
  const timezone = normalizeTimeZone_(spreadsheet.getSpreadsheetTimeZone());

  const sourceValues = sourceSheet.getDataRange().getValues();
  const regionValues = regionSheet.getDataRange().getValues();
  const goalValues = goalsSheet.getDataRange().getValues();
  const dashboard = buildDashboardData_(
    sourceValues,
    regionValues,
    goalValues,
    timezone,
    currentMonth,
    previousMonth
  );

  cache.put(cacheKey, JSON.stringify(dashboard), CONFIG.CACHE_SECONDS);
  return dashboard;
}

function configureSpreadsheet_() {
  requireDeploymentOwner_();
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  if (!spreadsheet) {
    throw new PlanoUserError('Execute esta função a partir do projeto vinculado à planilha.');
  }

  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', spreadsheet.getId());
  invalidateDashboardCache_();
  return `Planilha configurada: ${spreadsheet.getName()}`;
}

// Handler de gatilho por tempo (configurado manualmente no editor do Apps
// Script). Não é chamada pelo cliente via google.script.run — não existe
// nenhuma referência a ela em Index.html/PlanoAcaoClient.html — então não fica
// exposta ao navegador. Adicionar requireDeploymentOwner_ aqui quebraria a
// execução agendada, pois gatilhos por tempo nem sempre expõem
// Session.getActiveUser() da mesma forma que uma chamada manual no editor.
function atualizarDashboardDiariamente() {
  getRawDashboardData_(true, '', '');
  return { success: true };
}

function isDeploymentOwner_() {
  const activeEmail = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const effectiveEmail = String(Session.getEffectiveUser().getEmail() || '').trim().toLowerCase();
  return Boolean(activeEmail) && Boolean(effectiveEmail) && activeEmail === effectiveEmail;
}

function requireDeploymentOwner_() {
  if (!isDeploymentOwner_()) {
    throw new PlanoUserError('Esta função só pode ser executada pelo proprietário do projeto.');
  }
  return String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
}

function getSpreadsheet_() {
  const activeSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const configuredId = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  const spreadsheet = activeSpreadsheet || (configuredId
    ? SpreadsheetApp.openById(configuredId)
    : null);

  if (!spreadsheet) {
    throw new PlanoUserError(
      'Planilha não encontrada. Vincule o projeto à planilha ou configure a propriedade SPREADSHEET_ID.'
    );
  }

  return spreadsheet;
}

function getRequiredSheet_(spreadsheet, sheetName) {
  const sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) throw new PlanoUserError(`Aba obrigatória não encontrada: ${sheetName}`);
  return sheet;
}

function buildDashboardData_(
  sourceValues,
  regionValues,
  goalValues,
  timezone,
  selectedCurrentMonth,
  selectedPreviousMonth
) {
  timezone = normalizeTimeZone_(timezone);

  if (sourceValues.length < 2) {
    throw new PlanoUserError(`A aba ${CONFIG.SOURCE_SHEET} não possui dados.`);
  }

  const sourceHeaderRow = findHeaderRow_(sourceValues, ['data', 'cidade', 'base_geral']);
  const sourceHeaders = createHeaderMap_(sourceValues[sourceHeaderRow]);
  const dateIndex = getHeaderIndex_(sourceHeaders, 'data');
  const cityIndex = getHeaderIndex_(sourceHeaders, 'cidade');
  const baseIndex = getHeaderIndex_(sourceHeaders, 'base_geral');
  const cityMetadata = createCityMetadataMap_(regionValues);
  const valuesByMonth = new Map();

  sourceValues.slice(sourceHeaderRow + 1).forEach((row) => {
    const month = toMonthKey_(row[dateIndex]);
    const city = normalizeCity_(row[cityIndex]);
    const base = toNumber_(row[baseIndex]);

    if (!month || !city || base === null) return;

    if (!valuesByMonth.has(month)) valuesByMonth.set(month, new Map());
    const monthValues = valuesByMonth.get(month);
    const existing = monthValues.get(city) || { value: 0 };
    existing.value += base;
    monthValues.set(city, existing);
  });

  const months = Array.from(valuesByMonth.keys()).sort();
  if (months.length < 2) {
    throw new PlanoUserError('A base precisa ter pelo menos dois meses para comparação.');
  }

  const currentMonth = resolveMonth_(
    selectedCurrentMonth,
    months,
    defaultDashboardMonth_(months, timezone),
    'Mês analisado'
  );
  const currentMonthIndex = months.indexOf(currentMonth);
  const defaultPreviousMonth = currentMonthIndex > 0
    ? months[currentMonthIndex - 1]
    : months.find((month) => month !== currentMonth);
  const previousMonth = resolveMonth_(
    selectedPreviousMonth,
    months,
    defaultPreviousMonth,
    'Mês de comparação'
  );

  if (currentMonth === previousMonth) {
    throw new PlanoUserError('Selecione dois meses diferentes para realizar a comparação.');
  }

  const goalsByCity = createCityGoalsMap_(goalValues, currentMonth);
  const growthGoalsByCity = createCityGrowthGoalSeriesMap_(goalValues);

  const previousValues = valuesByMonth.get(previousMonth);
  const currentValues = valuesByMonth.get(currentMonth);
  const cities = new Set([...previousValues.keys(), ...currentValues.keys()]);

  // Somente os registros por cidade são calculados aqui. Resumo, rankings e a
  // lista de regionais dependem do conjunto de cidades visível a cada usuário
  // e são calculados uma única vez, em filterDashboardForUser_, a partir deste
  // "cities" completo — calculá-los aqui também seria trabalho descartado
  // (filterDashboardForUser_ sempre sobrescreve o resultado).
  const records = Array.from(cities).map((city) => {
    const previous = previousValues.get(city) || { value: 0 };
    const current = currentValues.get(city) || { value: 0 };
    const difference = current.value - previous.value;
    const metadata = cityMetadata.get(city) || {};
    const goals = goalsByCity.get(city) || createEmptyCityGoals_();

    return {
      city,
      regional: metadata.regional || 'Não informada',
      population: metadata.population || null,
      goals,
      previous: previous.value,
      current: current.value,
      difference,
      percentage: previous.value !== 0 ? difference / previous.value : 0,
      status: difference > 0 ? 'Crescimento' : difference < 0 ? 'Queda' : 'Estável',
    };
  });

  const orderedRecords = [...records].sort(
    (a, b) => b.difference - a.difference || a.city.localeCompare(b.city, 'pt-BR')
  );

  return {
    availableMonths: months
      .slice()
      .reverse()
      .map((month) => ({ key: month, label: formatMonthLabel_(month) })),
    period: {
      previousKey: previousMonth,
      currentKey: currentMonth,
      previousLabel: formatMonthLabel_(previousMonth),
      currentLabel: formatMonthLabel_(currentMonth),
    },
    cities: orderedRecords,
    cityHistories: createCityHistories_(
      valuesByMonth,
      months,
      cityMetadata,
      growthGoalsByCity
    ),
    updatedAt: Utilities.formatDate(new Date(), timezone, 'dd/MM/yyyy HH:mm:ss'),
  };
}

function createCityGoalsMap_(values, selectedMonth) {
  if (!values || !values.length) return new Map();

  const headerRow = findGoalsHeaderRow_(values, selectedMonth);
  const headers = values[headerRow];
  const cityIndex = findGoalCityHeaderIndex_(headers);
  const indicatorIndex = findGoalIndicatorHeaderIndex_(headers);
  const selectedMonthIndex = findGoalMonthHeaderIndex_(headers, selectedMonth);

  return createMatrixCityGoalsMap_(
    values.slice(headerRow + 1),
    cityIndex,
    indicatorIndex,
    selectedMonthIndex
  );
}

function createMatrixCityGoalsMap_(rows, cityIndex, indicatorIndex, monthIndex) {
  const goalsByCity = new Map();

  rows.forEach((row) => {
    const city = normalizeCity_(row[cityIndex]);
    const goalKey = normalizeGoalKey_(row[indicatorIndex]);
    const value = toNumber_(row[monthIndex]);
    if (!city || !goalKey || value === null || (goalKey !== 'growth' && value < 0)) return;

    const goals = goalsByCity.get(city) || createEmptyCityGoals_();
    goals[goalKey] = value;
    goalsByCity.set(city, goals);
  });

  return goalsByCity;
}

function createCityGrowthGoalSeriesMap_(values) {
  const seriesByCity = new Map();
  if (!values || !values.length) return seriesByCity;

  const headerRow = findGoalsHeaderRow_(values, '2000-01');
  const headers = values[headerRow];
  const cityIndex = findGoalCityHeaderIndex_(headers);
  const indicatorIndex = findGoalIndicatorHeaderIndex_(headers);
  const monthIndexes = createGoalMonthIndexes_(headers);

  values.slice(headerRow + 1).forEach((row) => {
    const city = normalizeCity_(row[cityIndex]);
    if (!city || normalizeGoalKey_(row[indicatorIndex]) !== 'growth') return;

    const series = {};
    monthIndexes.forEach((columnIndex, monthIndex) => {
      const value = toNumber_(row[columnIndex]);
      series[String(monthIndex + 1).padStart(2, '0')] = value;
    });
    seriesByCity.set(city, series);
  });

  return seriesByCity;
}

function createGoalMonthIndexes_(headers) {
  return Array.from({ length: 12 }, (unused, index) =>
    findGoalMonthHeaderIndex_(headers, `2000-${String(index + 1).padStart(2, '0')}`)
  );
}

function findGoalsHeaderRow_(values, selectedMonth) {
  const limit = Math.min(values.length, 10);
  for (let rowIndex = 0; rowIndex < limit; rowIndex += 1) {
    const headers = values[rowIndex];
    if (findGoalCityHeaderIndex_(headers, true) === null) continue;

    const isValidLayout = findGoalIndicatorHeaderIndex_(headers, true) !== null
      && findGoalMonthHeaderIndex_(headers, selectedMonth, true) !== null;
    if (isValidLayout) return rowIndex;
  }

  throw new PlanoUserError(
    `Formato inválido na aba ${CONFIG.GOALS_SHEET}. Verifique as colunas CIDADES, INDICADOR / SINALIZADOR e o mês selecionado.`
  );
}

function findGoalCityHeaderIndex_(headers, optional) {
  const normalizedHeaders = headers.map(normalizeHeader_);
  const preferredNames = ['cidade padronizada', 'cidade', 'municipio padronizado', 'municipio'];

  for (let index = 0; index < preferredNames.length; index += 1) {
    const headerIndex = normalizedHeaders.indexOf(preferredNames[index]);
    if (headerIndex >= 0) return headerIndex;
  }

  const fallbackIndex = normalizedHeaders.findIndex((header) =>
    header.includes('cidade') || header.includes('municipio')
  );
  if (fallbackIndex >= 0) return fallbackIndex;
  if (optional) return null;
  throw new PlanoUserError(`Coluna de cidade não encontrada na aba ${CONFIG.GOALS_SHEET}.`);
}

function findGoalIndicatorHeaderIndex_(headers, optional) {
  const index = headers.findIndex((header) => {
    const normalized = normalizeHeader_(header);
    return normalized.includes('indicador') || normalized.includes('sinalizador');
  });
  if (index >= 0) return index;
  if (optional) return null;
  throw new PlanoUserError(`Coluna INDICADOR / SINALIZADOR não encontrada na aba ${CONFIG.GOALS_SHEET}.`);
}

function findGoalMonthHeaderIndex_(headers, selectedMonth, optional) {
  const monthNumber = Number(String(selectedMonth || '').slice(5, 7));
  const monthHeaders = ['jan', 'fev', 'mar', 'abr', 'mai', 'jun', 'jul', 'ago', 'set', 'out', 'nov', 'dez'];
  const expectedHeader = monthHeaders[monthNumber - 1] || '';
  const index = headers.findIndex((header) => normalizeHeader_(header) === expectedHeader);
  if (index >= 0) return index;
  if (optional) return null;
  throw new PlanoUserError(`Coluna do mês ${expectedHeader.toUpperCase()} não encontrada na aba ${CONFIG.GOALS_SHEET}.`);
}

function createEmptyCityGoals_() {
  return { budget: null, effective: null, installation: null, growth: null };
}

function normalizeGoalKey_(value) {
  const normalized = normalizeHeader_(value);
  if (normalized === 'budget' || normalized.includes('orcament')) return 'budget';
  if (normalized === 'effective' || normalized.includes('efetiv')) return 'effective';
  if (normalized === 'installation' || normalized.includes('instala')) return 'installation';
  if (normalized === 'growth' || normalized.includes('crescimento')) return 'growth';
  return '';
}

function getDashboardCacheRevision_() {
  const properties = PropertiesService.getScriptProperties();
  let revision = properties.getProperty(CONFIG.CACHE_REVISION_PROPERTY);
  if (!revision) {
    revision = Utilities.getUuid();
    properties.setProperty(CONFIG.CACHE_REVISION_PROPERTY, revision);
  }
  return revision;
}

function invalidateDashboardCache_() {
  const revision = Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperty(CONFIG.CACHE_REVISION_PROPERTY, revision);
  return revision;
}

function createCacheKey_(revision, currentMonth, previousMonth) {
  return [
    CONFIG.CACHE_KEY,
    revision,
    currentMonth || 'latest',
    previousMonth || 'previous',
  ].join(':');
}

// summary/rankings NÃO fazem parte do retorno: o cliente sempre recalcula os
// dois a partir de "cities" (createClientSummary em Index.html), porque
// precisa refazer as contas depois de aplicar a exclusão manual de cidades
// (state.excludedCities), uma preferência que só existe no navegador — o
// servidor não tem como sabe quais cidades o usuário excluiu. Calcular aqui
// também seria trabalho sempre descartado pelo cliente.
function filterDashboardForUser_(dashboard, user) {
  const cities = dashboard.cities.filter((record) =>
    hasAppRegionalAccess_(user, record.regional)
  );

  return Object.assign({}, dashboard, {
    regions: Array.from(new Set(cities.map((record) => record.regional))).sort((a, b) =>
      a.localeCompare(b, 'pt-BR')
    ),
    cities,
    cityHistories: dashboard.cityHistories.filter((history) =>
      hasAppRegionalAccess_(user, history.regional)
    ),
  });
}

// Mês usado como "atual" quando o usuário não escolheu nenhum: o último mês
// que já fechou de verdade, não o mês corrente em andamento. Um mês em curso
// (ex.: setembro visto no dia 3) tem base parcial e comparar contra ele
// distorceria o ranking de crescimento/queda logo no início de cada mês. O
// usuário sempre pode trocar manualmente para o mês atual no seletor.
function defaultDashboardMonth_(months, timezone) {
  const currentCalendarMonth = Utilities.formatDate(new Date(), timezone, 'yyyy-MM');
  const closedMonths = months.filter((month) => month < currentCalendarMonth);
  return closedMonths.length
    ? closedMonths[closedMonths.length - 1]
    : months[months.length - 1];
}

function resolveMonth_(selectedMonth, availableMonths, fallbackMonth, fieldName) {
  const month = selectedMonth || fallbackMonth;
  if (!availableMonths.includes(month)) {
    throw new PlanoUserError(`${fieldName} não encontrado na base: ${month}`);
  }
  return month;
}

function createCityHistories_(valuesByMonth, months, cityMetadata, growthGoalsByCity) {
  const cities = new Set();
  valuesByMonth.forEach((monthValues) => {
    monthValues.forEach((unused, city) => cities.add(city));
  });

  return Array.from(cities)
    .sort((a, b) => a.localeCompare(b, 'pt-BR'))
    .map((city) => ({
      city,
      regional: (cityMetadata.get(city) || {}).regional || 'Não informada',
      growthGoals: growthGoalsByCity.get(city) || {},
      points: months.map((month) => {
        const record = valuesByMonth.get(month).get(city);
        return {
          monthKey: month,
          label: formatMonthLabel_(month),
          value: record ? record.value : null,
        };
      }),
    }));
}

function createCityMetadataMap_(values) {
  if (!values.length) return new Map();

  const headerRow = findHeaderRow_(values, ['cidade', 'regional']);
  const headerValues = values[headerRow];
  const headers = createHeaderMap_(headerValues);
  const cityIndex = getHeaderIndex_(headers, 'cidade');
  const regionalIndex = getHeaderIndex_(headers, 'regional');
  // A aba legada já possui outra coluna chamada "População" mais à direita.
  // A coluna nova é a primeira ocorrência; o Map de cabeçalhos preserva apenas
  // a última, portanto a busca precisa ser feita diretamente na linha original.
  const populationIndex = findFirstHeaderIndex_(headerValues, 'populacao');
  const metadata = new Map();

  values.slice(headerRow + 1).forEach((row) => {
    const city = normalizeCity_(row[cityIndex]);
    const regional = String(row[regionalIndex] || '').trim();
    if (!city) return;

    const populationValue = populationIndex === null ? null : toNumber_(row[populationIndex]);
    const population = isPositiveNumber_(populationValue) ? populationValue : null;
    const existing = metadata.get(city) || {};
    metadata.set(city, {
      regional: regional || existing.regional || '',
      population: population || existing.population || null,
    });
  });

  return metadata;
}


function createHeaderMap_(headers) {
  return headers.reduce((map, header, index) => {
    map.set(normalizeHeader_(header), index);
    return map;
  }, new Map());
}

function findHeaderRow_(values, requiredHeaders) {
  const normalizedRequiredHeaders = requiredHeaders.map(normalizeHeader_);
  const limit = Math.min(values.length, 10);

  for (let rowIndex = 0; rowIndex < limit; rowIndex += 1) {
    const headers = createHeaderMap_(values[rowIndex]);
    if (normalizedRequiredHeaders.every((header) => headers.has(header))) return rowIndex;
  }

  throw new PlanoUserError(`Cabeçalhos obrigatórios não encontrados: ${requiredHeaders.join(', ')}`);
}

function getHeaderIndex_(headers, name) {
  const index = headers.get(normalizeHeader_(name));
  if (index === undefined) throw new PlanoUserError(`Coluna obrigatória não encontrada: ${name}`);
  return index;
}

function findFirstHeaderIndex_(headers, name) {
  const normalizedName = normalizeHeader_(name);
  const index = headers.findIndex((header) => normalizeHeader_(header) === normalizedName);
  return index < 0 ? null : index;
}

function isPositiveNumber_(value) {
  return typeof value === 'number' && Number.isFinite(value) && value > 0;
}

function normalizeHeader_(value) {
  return String(value || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .trim()
    .toLowerCase();
}

function normalizeCity_(value) {
  return String(value || '')
    .trim()
    .replace(/\s*\/\s*/g, '/')
    .toUpperCase();
}

function toNumber_(value) {
  if (typeof value === 'number' && Number.isFinite(value)) return value;

  const normalized = String(value || '')
    .trim()
    .replace(/\./g, '')
    .replace(',', '.');
  if (!normalized) return null;

  const number = Number(normalized);
  return Number.isFinite(number) ? number : null;
}

function toMonthKey_(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(value, 'UTC', 'yyyy-MM');
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    const date = new Date(Math.round((value - 25569) * 86400000));
    return Utilities.formatDate(date, 'UTC', 'yyyy-MM');
  }

  const text = String(value || '').trim();
  let match = text.match(/^(\d{4})-(\d{2})(?:-\d{2})?/);
  if (match) return `${match[1]}-${match[2]}`;

  match = text.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (match) return `${match[3]}-${match[2]}`;

  return '';
}

function normalizeMonthKey_(value) {
  const month = String(value || '').trim();
  return /^\d{4}-(0[1-9]|1[0-2])$/.test(month) ? month : '';
}

function normalizeTimeZone_(value) {
  return typeof value === 'string' && value.trim() ? value.trim() : 'America/Fortaleza';
}

function formatMonthLabel_(monthKey) {
  const [year, month] = monthKey.split('-').map(Number);
  const months = [
    'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
    'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro',
  ];
  return `${months[month - 1]} de ${year}`;
}


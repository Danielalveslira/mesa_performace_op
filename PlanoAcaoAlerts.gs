const PLANO_ALERT_HANDLER = 'processarAlertasPlanos_';
const PLANO_ALERT_DEFAULTS = Object.freeze([
  Object.freeze({ key: 'ativo', value: 'TRUE', description: 'Ativa ou desativa o envio automático.' }),
  Object.freeze({ key: 'dias_antecedencia_prazo', value: '2', description: 'Envia um aviso quando faltarem até esta quantidade de dias para o prazo.' }),
  Object.freeze({ key: 'dias_sem_atualizacao', value: '7', description: 'Quantidade mínima de dias sem andamento para gerar alerta.' }),
  Object.freeze({ key: 'intervalo_atraso_dias', value: '3', description: 'Intervalo entre novos avisos de plano atrasado.' }),
  Object.freeze({ key: 'intervalo_sem_atualizacao_dias', value: '7', description: 'Intervalo entre novos avisos de plano sem atualização.' }),
  Object.freeze({ key: 'escalar_atraso_dias', value: '3', description: 'Após esta quantidade de dias de atraso, inclui gestores da regional.' }),
  Object.freeze({ key: 'escalar_sem_atualizacao_dias', value: '14', description: 'Após esta quantidade de dias sem andamento, inclui gestores da regional.' }),
  Object.freeze({ key: 'alerta_base_piorou', value: 'TRUE', description: 'Alerta quando um novo mês fica abaixo da base registrada na criação.' }),
  Object.freeze({ key: 'hora_execucao', value: '8', description: 'Hora aproximada da execução diária. Reexecute setupPlanoAlertas após alterar.' }),
  Object.freeze({ key: 'nome_remetente', value: 'Mesa Operacional', description: 'Nome exibido como remetente dos e-mails.' }),
]);

function setupPlanoAlertas() {
  requireDeploymentOwner_();
  setupPlanoAcao_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    seedPlanoAlertConfig_();
    const config = getPlanoAlertConfig_();
    ScriptApp.getProjectTriggers()
      .filter((trigger) => trigger.getHandlerFunction() === PLANO_ALERT_HANDLER)
      .forEach((trigger) => ScriptApp.deleteTrigger(trigger));

    ScriptApp.newTrigger(PLANO_ALERT_HANDLER)
      .timeBased()
      .everyDays(1)
      .atHour(config.executionHour)
      .inTimezone(getPlanoTimeZone_())
      .create();

    return {
      success: true,
      message: `Alertas configurados para execução diária por volta de ${String(config.executionHour).padStart(2, '0')}:00.`,
      sheets: [PLANO_ACAO_CONFIG.SHEETS.ALERT_CONFIG, PLANO_ACAO_CONFIG.SHEETS.ALERT_HISTORY],
    };
  } finally {
    lock.releaseLock();
  }
}

function previewPlanoAlerts() {
  return planoRunPublic_('previewPlanoAlerts', () => {
    const user = getAppUser_();
    requirePlanoPermission_(user, 'administer');
    return runPlanoAlerts_(true, false);
  });
}

function executarAlertasAgora() {
  return planoRunPublic_('executarAlertasAgora', () => {
    requireDeploymentOwner_();
    return runPlanoAlerts_(false, true);
  });
}

function processarAlertasPlanos_() {
  return runPlanoAlerts_(false, true);
}

function runPlanoAlerts_(dryRun, forceDashboardRefresh) {
  // Domínio próprio ('plano-alertas'): processar alertas só lê planos e
  // grava em Alertas_Historico, não compete pelo domínio 'plano-crud'
  // (Planos_Acao/Plano_*).
  const lock = acquirePlanoNamedLock_('plano-alertas', 5000);
  if (!lock) {
    return { success: false, skipped: true, message: 'Já existe outra execução de alertas em andamento.' };
  }

  try {
    const config = getPlanoAlertConfig_();
    if (!config.active) {
      return { success: true, disabled: true, message: 'Alertas desativados em Alertas_Config.' };
    }

    const timezone = getPlanoTimeZone_();
    const now = new Date();
    const users = readPlanoAlertUsers_();
    const plans = readPlanoRecords_(PLANO_ACAO_CONFIG.SHEETS.PLANS)
      .filter((plan) => toPlanoBoolean_(plan.ativo, true))
      .filter((plan) => !['concluido', 'cancelado'].includes(String(plan.status || '')));
    const histories = readPlanoAlertCityHistories_(forceDashboardRefresh);
    const alerts = plans.flatMap((plan) => evaluatePlanoAlerts_(plan, config, histories, now, timezone));
    const historyState = readPlanoAlertHistoryState_(now, timezone);
    const queue = createPlanoAlertQueue_(alerts, users, config, historyState);

    if (dryRun) {
      return createPlanoAlertPreview_(plans.length, alerts, queue, config, now, timezone);
    }

    const result = sendPlanoAlertQueue_(queue.pending, config, now);
    queue.unresolved.forEach((alert) => result.historyRecords.push(createPlanoAlertHistoryRecord_(
      alert,
      '',
      'erro',
      'Nenhum usuário ativo e autorizado foi encontrado para receber o alerta.',
      now
    )));
    // Grava todo o histórico de alertas desta execução em uma única chamada em
    // lote, em vez de uma escrita por destinatário/alerta.
    appendPlanoRecords_(getPlanoSheet_(PLANO_ACAO_CONFIG.SHEETS.ALERT_HISTORY), result.historyRecords);
    SpreadsheetApp.flush();

    return {
      success: result.failedEmails === 0 && queue.unresolved.length === 0,
      activePlans: plans.length,
      detectedAlerts: alerts.length,
      pendingDeliveries: queue.pending.length,
      suppressedDeliveries: queue.suppressed,
      recentAttempts: queue.deferred,
      unresolvedAlerts: queue.unresolved.length,
      sentEmails: result.sentEmails,
      sentAlerts: result.sentAlerts,
      failedEmails: result.failedEmails,
    };
  } finally {
    lock.release();
  }
}

function seedPlanoAlertConfig_() {
  const sheet = getPlanoSheet_(PLANO_ACAO_CONFIG.SHEETS.ALERT_CONFIG);
  const existing = new Set(
    readPlanoRecords_(PLANO_ACAO_CONFIG.SHEETS.ALERT_CONFIG)
      .map((record) => String(record.chave || '').trim().toLowerCase())
  );
  const now = new Date();
  PLANO_ALERT_DEFAULTS.forEach((item) => {
    if (existing.has(item.key)) return;
    appendPlanoRecord_(sheet, {
      chave: item.key,
      valor: item.value,
      descricao: item.description,
      atualizado_em: now,
    });
  });
  SpreadsheetApp.flush();
}

function getPlanoAlertConfig_() {
  const values = new Map(
    readPlanoRecords_(PLANO_ACAO_CONFIG.SHEETS.ALERT_CONFIG)
      .map((record) => [String(record.chave || '').trim().toLowerCase(), record.valor])
  );
  const defaultValue = (key) => PLANO_ALERT_DEFAULTS.find((item) => item.key === key).value;
  const value = (key) => values.has(key) ? values.get(key) : defaultValue(key);
  return {
    active: toPlanoBoolean_(value('ativo'), true),
    deadlineLeadDays: planoAlertInteger_(value('dias_antecedencia_prazo'), 2, 0, 30),
    staleDays: planoAlertInteger_(value('dias_sem_atualizacao'), 7, 1, 365),
    overdueIntervalDays: planoAlertInteger_(value('intervalo_atraso_dias'), 3, 1, 365),
    staleIntervalDays: planoAlertInteger_(value('intervalo_sem_atualizacao_dias'), 7, 1, 365),
    escalateOverdueDays: planoAlertInteger_(value('escalar_atraso_dias'), 3, 1, 365),
    escalateStaleDays: planoAlertInteger_(value('escalar_sem_atualizacao_dias'), 14, 1, 365),
    baseDeclineEnabled: toPlanoBoolean_(value('alerta_base_piorou'), true),
    executionHour: planoAlertInteger_(value('hora_execucao'), 8, 0, 23),
    senderName: String(value('nome_remetente') || 'Mesa Operacional').trim().slice(0, 80),
  };
}

/**
 * Mesma leitura de getPlanoAlertConfig_().staleDays, mas tolerante a
 * Alertas_Config ainda não configurada (ex.: setupPlanoAlertas nunca
 * executado) — usada pelo bootstrap de planos, que não pode falhar por
 * causa de um módulo opcional.
 */
function getPlanoAlertStaleDaysSafe_() {
  try {
    return getPlanoAlertConfig_().staleDays;
  } catch (error) {
    const fallback = PLANO_ALERT_DEFAULTS.find((item) => item.key === 'dias_sem_atualizacao');
    return Number(fallback.value);
  }
}

function readPlanoAlertUsers_() {
  return readPlanoRecords_(PLANO_ACAO_CONFIG.SHEETS.USERS)
    .filter((record) => toPlanoBoolean_(record.ativo, false))
    .map(createPlanoUserFromRecord_)
    .filter(Boolean);
}

function readPlanoAlertCityHistories_(forceRefresh) {
  try {
    const dashboard = getRawDashboardData_(Boolean(forceRefresh), '', '');
    return new Map((dashboard.cityHistories || []).map((history) => [normalizeCity_(history.city), history]));
  } catch (error) {
    console.warn(`Os alertas de evolução da base foram ignorados: ${error && error.message ? error.message : error}`);
    return new Map();
  }
}

function evaluatePlanoAlerts_(plan, config, histories, now, timezone) {
  const alerts = [];
  const todayKey = Utilities.formatDate(now, timezone, 'yyyy-MM-dd');
  const todayDay = planoAlertDayNumber_(todayKey);
  const deadlineKey = planoAlertDateKey_(plan.prazo, timezone);

  if (deadlineKey) {
    const daysUntil = planoAlertDayNumber_(deadlineKey) - todayDay;
    if (daysUntil >= 0 && daysUntil <= config.deadlineLeadDays) {
      alerts.push(createPlanoAlert_(plan, 'prazo_proximo', deadlineKey, {
        severity: 'atencao',
        message: daysUntil === 0 ? 'O prazo vence hoje.' : `O prazo vence em ${daysUntil} ${daysUntil === 1 ? 'dia' : 'dias'}.`,
        deadline: deadlineKey,
        days: daysUntil,
      }));
    } else if (daysUntil < 0) {
      const daysOverdue = Math.abs(daysUntil);
      const cycle = Math.floor((daysOverdue - 1) / config.overdueIntervalDays);
      alerts.push(createPlanoAlert_(plan, 'atrasado', `${deadlineKey}:ciclo-${cycle}`, {
        severity: 'critico',
        message: `O plano está atrasado há ${daysOverdue} ${daysOverdue === 1 ? 'dia' : 'dias'}.`,
        deadline: deadlineKey,
        days: daysOverdue,
      }));
    }
  }

  const updatedAt = new Date(plan.atualizado_em || plan.criado_em || 0);
  if (!Number.isNaN(updatedAt.getTime())) {
    const staleDays = Math.max(0, Math.floor((now.getTime() - updatedAt.getTime()) / 86400000));
    if (staleDays >= config.staleDays) {
      const cycle = Math.floor((staleDays - config.staleDays) / config.staleIntervalDays);
      alerts.push(createPlanoAlert_(plan, 'sem_atualizacao', `ciclo-${cycle}`, {
        severity: 'atencao',
        message: `O plano está há ${staleDays} dias sem registro de andamento.`,
        days: staleDays,
      }));
    }
  }

  if (config.baseDeclineEnabled) {
    const decline = createPlanoBaseDeclineAlert_(plan, histories);
    if (decline) alerts.push(decline);
  }
  return alerts;
}

function createPlanoBaseDeclineAlert_(plan, histories) {
  const baselineRaw = plan.base_atual;
  if (baselineRaw === '' || baselineRaw === null || baselineRaw === undefined) return null;
  const baseline = Number(baselineRaw);
  if (!Number.isFinite(baseline)) return null;
  const history = histories.get(normalizeCity_(plan.cidade));
  if (!history || !Array.isArray(history.points)) return null;
  const analyzedMonth = planoAlertMonthKey_(plan.periodo_analisado);
  const points = history.points
    .filter((point) => point.value !== '' && point.value !== null && point.value !== undefined)
    .filter((point) => Number.isFinite(Number(point.value)))
    .filter((point) => !analyzedMonth || planoAlertMonthKey_(point.monthKey) > analyzedMonth)
    .sort((a, b) => String(a.monthKey || '').localeCompare(String(b.monthKey || '')));
  const latest = points[points.length - 1];
  if (!latest || Number(latest.value) >= baseline) return null;
  const current = Number(latest.value);
  const difference = current - baseline;
  const percentage = baseline ? difference / baseline : null;
  return createPlanoAlert_(plan, 'base_piorou', planoAlertMonthKey_(latest.monthKey), {
    severity: 'critico',
    message: `A base caiu ${planoAlertSignedInteger_(difference)}${percentage === null ? '' : ` (${planoAlertPercent_(percentage)})`} desde a criação do plano.`,
    baseAtCreation: baseline,
    currentBase: current,
    difference,
    monthLabel: latest.label || planoAlertMonthKey_(latest.monthKey),
  });
}

function createPlanoAlert_(plan, type, reference, details) {
  return Object.assign({
    planId: String(plan.id || ''),
    type,
    reference: String(reference || ''),
    key: `${String(plan.id || '')}|${type}|${String(reference || '')}`,
    city: String(plan.cidade || ''),
    regional: String(plan.regional || ''),
    action: String(plan.o_que || ''),
    responsible: String(plan.responsavel || ''),
    responsibleEmail: String(plan.responsavel_email || '').trim().toLowerCase(),
    createdBy: String(plan.criado_por || '').trim().toLowerCase(),
    status: String(plan.status || ''),
  }, details || {});
}

function createPlanoAlertQueue_(alerts, users, config, historyState) {
  const pending = [];
  const unresolved = [];
  let suppressed = 0;
  let deferred = 0;

  alerts.forEach((alert) => {
    const recipients = resolvePlanoAlertRecipients_(alert, users, config);
    if (!recipients.length) {
      if (historyState.attemptedTodayKeys.has(`${alert.key}|`)) {
        deferred += 1;
        return;
      }
      unresolved.push(alert);
      return;
    }
    recipients.forEach((recipient) => {
      const deliveryKey = `${alert.key}|${recipient.email}`;
      if (historyState.sentKeys.has(deliveryKey)) {
        suppressed += 1;
        return;
      }
      if (historyState.attemptedTodayKeys.has(deliveryKey)) {
        deferred += 1;
        return;
      }
      pending.push({ alert, recipient, deliveryKey });
    });
  });
  return { pending, unresolved, suppressed, deferred };
}

function resolvePlanoAlertRecipients_(alert, users, config) {
  const recipients = new Map();
  const canAccess = (user) => user && hasAppRegionalAccess_(user, alert.regional);
  const add = (user) => {
    if (canAccess(user)) recipients.set(user.email, user);
  };

  add(users.find((user) => user.email === alert.responsibleEmail));
  if (!recipients.size && alert.responsible) {
    const responsibleName = normalizeHeader_(alert.responsible);
    const matches = users.filter((user) => normalizeHeader_(user.name) === responsibleName && canAccess(user));
    if (matches.length === 1) add(matches[0]);
  }
  if (!recipients.size) add(users.find((user) => user.email === alert.createdBy));

  const escalate = alert.type === 'atrasado' && Number(alert.days) >= config.escalateOverdueDays
    || alert.type === 'sem_atualizacao' && Number(alert.days) >= config.escalateStaleDays;
  if (escalate || !recipients.size) {
    users
      .filter((user) => user.receivesManagementAlerts || user.role === 'administrador')
      .forEach(add);
  }
  return Array.from(recipients.values());
}

function readPlanoAlertHistoryState_(now, timezone) {
  const todayKey = Utilities.formatDate(now, timezone, 'yyyy-MM-dd');
  const sentKeys = new Set();
  const attemptedTodayKeys = new Set();
  readPlanoRecords_(PLANO_ACAO_CONFIG.SHEETS.ALERT_HISTORY).forEach((record) => {
    const key = String(record.chave_unica || '').trim();
    if (!key) return;
    if (String(record.status || '').trim().toLowerCase() === 'enviado') sentKeys.add(key);
    if (planoAlertDateKey_(record.enviado_em, timezone) === todayKey) attemptedTodayKeys.add(key);
  });
  return { sentKeys, attemptedTodayKeys };
}

function sendPlanoAlertQueue_(pending, config, now) {
  const groups = new Map();
  pending.forEach((delivery) => {
    if (!groups.has(delivery.recipient.email)) groups.set(delivery.recipient.email, []);
    groups.get(delivery.recipient.email).push(delivery);
  });

  let remainingQuota = MailApp.getRemainingDailyQuota();
  let sentEmails = 0;
  let sentAlerts = 0;
  let failedEmails = 0;
  const historyRecords = [];

  Array.from(groups.entries()).forEach(([email, deliveries]) => {
    const subject = `[Mesa Operacional] ${deliveries.length} ${deliveries.length === 1 ? 'plano precisa' : 'planos precisam'} de atenção`;
    if (remainingQuota <= 0) {
      failedEmails += 1;
      deliveries.forEach((delivery) => historyRecords.push(
        createPlanoAlertHistoryRecord_(delivery.alert, email, 'erro', 'Cota diária de e-mails esgotada.', now, subject)
      ));
      return;
    }
    try {
      MailApp.sendEmail({
        to: email,
        subject,
        body: createPlanoAlertTextBody_(deliveries),
        htmlBody: createPlanoAlertHtmlBody_(deliveries),
        name: config.senderName,
      });
      remainingQuota -= 1;
      sentEmails += 1;
      sentAlerts += deliveries.length;
      deliveries.forEach((delivery) => historyRecords.push(
        createPlanoAlertHistoryRecord_(delivery.alert, email, 'enviado', '', now, subject)
      ));
    } catch (error) {
      failedEmails += 1;
      const message = error && error.message ? error.message : String(error);
      deliveries.forEach((delivery) => historyRecords.push(
        createPlanoAlertHistoryRecord_(delivery.alert, email, 'erro', message, now, subject)
      ));
    }
  });
  return { sentEmails, sentAlerts, failedEmails, historyRecords };
}

function createPlanoAlertTextBody_(deliveries) {
  const lines = ['Mesa Operacional', '', 'Planos que precisam de atenção:', ''];
  deliveries.forEach((delivery, index) => {
    const alert = delivery.alert;
    lines.push(`${index + 1}. ${alert.city} — ${planoAlertTypeLabel_(alert.type)}`);
    lines.push(alert.message);
    lines.push(`Regional: ${alert.regional || 'Não informada'} | Responsável: ${alert.responsible || 'Não informado'}`);
    lines.push('');
  });
  lines.push('Acesse a Mesa Operacional para revisar e registrar o andamento.');
  return lines.join('\n');
}

function createPlanoAlertHtmlBody_(deliveries) {
  const items = deliveries.map((delivery) => {
    const alert = delivery.alert;
    return `<tr><td style="padding:12px;border-bottom:1px solid #d0d7de"><strong>${planoAlertEscapeHtml_(alert.city)}</strong><br><span style="color:#57606a">${planoAlertEscapeHtml_(alert.regional || 'Regional não informada')}</span></td><td style="padding:12px;border-bottom:1px solid #d0d7de"><strong>${planoAlertEscapeHtml_(planoAlertTypeLabel_(alert.type))}</strong><br>${planoAlertEscapeHtml_(alert.message)}</td><td style="padding:12px;border-bottom:1px solid #d0d7de">${planoAlertEscapeHtml_(alert.responsible || 'Não informado')}</td></tr>`;
  }).join('');
  const appUrl = String(ScriptApp.getService().getUrl() || '').trim();
  const action = appUrl
    ? `<p style="margin:20px 0 0"><a href="${planoAlertEscapeHtml_(appUrl)}" style="display:inline-block;background:#0969da;color:#fff;text-decoration:none;border-radius:6px;padding:10px 14px;font-weight:600">Abrir Mesa Operacional</a></p>`
    : '<p style="margin:20px 0 0;color:#57606a">Acesse a Mesa Operacional para registrar o andamento.</p>';
  return `<div style="font-family:Arial,sans-serif;color:#1f2328;max-width:760px"><h2 style="margin:0 0 6px">Mesa Operacional</h2><p style="margin:0 0 18px;color:#57606a">Resumo automático dos planos que precisam de atenção.</p><table style="width:100%;border-collapse:collapse;border:1px solid #d0d7de"><thead><tr style="background:#f6f8fa"><th style="padding:10px;text-align:left">Cidade</th><th style="padding:10px;text-align:left">Alerta</th><th style="padding:10px;text-align:left">Responsável</th></tr></thead><tbody>${items}</tbody></table>${action}<p style="margin-top:18px;color:#6e7781;font-size:12px">Mensagem automática. Atualize o plano para interromper alertas de atraso ou falta de andamento.</p></div>`;
}

function createPlanoAlertHistoryRecord_(alert, recipient, status, error, sentAt, subject) {
  return {
    id: Utilities.getUuid(),
    plano_id: alert.planId,
    tipo: alert.type,
    referencia: alert.reference,
    chave_unica: `${alert.key}|${String(recipient || '').trim().toLowerCase()}`,
    destinatario: String(recipient || '').trim().toLowerCase(),
    assunto: subject || `[Mesa Operacional] ${planoAlertTypeLabel_(alert.type)}`,
    enviado_em: sentAt || new Date(),
    status,
    erro: String(error || '').slice(0, 1000),
  };
}

function createPlanoAlertPreview_(activePlans, alerts, queue, config, now, timezone) {
  const grouped = new Map();
  queue.pending.forEach((delivery) => {
    if (!grouped.has(delivery.alert.key)) {
      grouped.set(delivery.alert.key, Object.assign({}, delivery.alert, { recipients: [] }));
    }
    grouped.get(delivery.alert.key).recipients.push(delivery.recipient.email);
  });
  queue.unresolved.forEach((alert) => {
    grouped.set(alert.key, Object.assign({}, alert, { recipients: [] }));
  });
  return {
    success: true,
    dryRun: true,
    activePlans,
    detectedAlerts: alerts.length,
    pendingAlerts: grouped.size,
    pendingDeliveries: queue.pending.length,
    suppressedDeliveries: queue.suppressed,
    recentAttempts: queue.deferred,
    unresolvedAlerts: queue.unresolved.length,
    config: {
      deadlineLeadDays: config.deadlineLeadDays,
      staleDays: config.staleDays,
      executionHour: config.executionHour,
    },
    // Instante absoluto (ISO) da próxima execução agendada do gatilho diário.
    // É um timestamp absoluto, não um horário "de parede" — o navegador pode
    // comparar direto com o próprio relógio sem se preocupar com fuso horário.
    nextScheduledRunAt: computePlanoNextAlertRun_(config, now, timezone).toISOString(),
    items: Array.from(grouped.values()).slice(0, 100).map((alert) => ({
      city: alert.city,
      regional: alert.regional,
      type: alert.type,
      typeLabel: planoAlertTypeLabel_(alert.type),
      message: alert.message,
      recipients: alert.recipients,
    })),
  };
}

function computePlanoNextAlertRun_(config, now, timezone) {
  const todayKey = Utilities.formatDate(now, timezone, 'yyyy-MM-dd');
  const hour = String(config.executionHour).padStart(2, '0');
  const todayRun = Utilities.parseDate(`${todayKey} ${hour}:00:00`, timezone, 'yyyy-MM-dd HH:mm:ss');
  if (todayRun.getTime() > now.getTime()) return todayRun;
  // Já passou do horário de hoje: a próxima execução é amanhã no mesmo
  // horário. America/Fortaleza não observa horário de verão, então somar
  // 24h em milissegundos é sempre exato.
  return new Date(todayRun.getTime() + 24 * 60 * 60 * 1000);
}

function planoAlertTypeLabel_(type) {
  return {
    prazo_proximo: 'Prazo próximo',
    atrasado: 'Plano atrasado',
    sem_atualizacao: 'Sem atualização',
    base_piorou: 'Base piorou',
  }[type] || type;
}

function planoAlertDateKey_(value, timezone) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(value, timezone, 'yyyy-MM-dd');
  }
  const match = String(value || '').match(/^(\d{4})-(\d{2})-(\d{2})/);
  return match ? `${match[1]}-${match[2]}-${match[3]}` : '';
}

function planoAlertDayNumber_(dateKey) {
  const match = String(dateKey || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  return match ? Math.floor(Date.UTC(Number(match[1]), Number(match[2]) - 1, Number(match[3])) / 86400000) : 0;
}

function planoAlertMonthKey_(value) {
  const match = String(value || '').match(/^(\d{4})-(\d{2})/);
  return match ? `${match[1]}-${match[2]}` : '';
}

function planoAlertInteger_(value, fallback, minimum, maximum) {
  const number = Math.trunc(Number(value));
  if (!Number.isFinite(number)) return fallback;
  return Math.max(minimum, Math.min(maximum, number));
}

function planoAlertSignedInteger_(value) {
  const number = Number(value) || 0;
  return `${number > 0 ? '+' : ''}${number.toLocaleString('pt-BR', { maximumFractionDigits: 0 })}`;
}

function planoAlertPercent_(value) {
  return Number(value || 0).toLocaleString('pt-BR', { style: 'percent', minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function planoAlertEscapeHtml_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

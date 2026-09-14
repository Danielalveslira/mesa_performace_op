<script>
  const planoState = {
    available: false,
    user: null,
    permissions: {},
    plans: [],
    options: { statuses: [], priorities: [], responsibles: [] },
    editingPlan: null,
    selectedDetail: null,
    selectedCity: '',
    attentionFilter: '',
    detailCache: new Map(),
    detailRequestId: 0,
    formOperationId: '',
    updateOperationId: '',
  };
  const PLANO_MAX_IMAGE_BYTES = 4 * 1024 * 1024;
  const PLANO_MAX_IMAGES_PER_UPLOAD = 3;
  const PLANO_STALE_DAYS = 7;
  const PLANO_ALLOWED_IMAGE_TYPES = ['image/jpeg', 'image/png', 'image/webp'];
  let planoAlertCountdownTimer = null;
  let planoLastAlertSendResult = null;
  let planoLastAlertPendingDeliveries = 0;

  window.__activeAppView = 'dashboard';

  document.getElementById('nav-dashboard').addEventListener('click', () => planoShowView('dashboard'));
  document.getElementById('nav-plans').addEventListener('click', () => planoShowView('plans'));
  document.getElementById('new-plan-button').addEventListener('click', () => planoOpenForm());
  document.getElementById('plan-search').addEventListener('input', planoRenderList);
  document.getElementById('plan-status-filter').addEventListener('change', planoRenderList);
  document.getElementById('plan-region-filter').addEventListener('change', planoRenderList);
  document.getElementById('plan-priority-filter').addEventListener('change', planoRenderList);
  document.getElementById('clear-attention-filter').addEventListener('click', () => planoSetAttentionFilter(''));
  document.getElementById('preview-alerts-button').addEventListener('click', planoPreviewAlerts);
  document.getElementById('send-alerts-now-button').addEventListener('click', planoSendAlertsNow);
  document.getElementById('plan-city').addEventListener('change', planoSyncPlanRegion);
  document.getElementById('plan-responsible-email').addEventListener('change', planoSyncResponsibleName);
  document.getElementById('plan-form').addEventListener('submit', planoSubmitForm);
  document.getElementById('plan-update-form').addEventListener('submit', planoSubmitUpdate);

  document.querySelectorAll('[data-close-dialog]').forEach((button) => {
    button.addEventListener('click', () => document.getElementById(button.dataset.closeDialog).close());
  });

  ['plan-form-dialog', 'plan-detail-dialog', 'plan-update-dialog', 'plan-image-dialog', 'alert-preview-dialog'].forEach((id) => {
    document.getElementById(id).addEventListener('click', (event) => {
      if (event.target === event.currentTarget) event.currentTarget.close();
    });
  });
  document.getElementById('plan-image-dialog').addEventListener('close', () => {
    const image = document.getElementById('plan-image-preview');
    image.removeAttribute('src');
    image.alt = '';
  });
  document.getElementById('alert-preview-dialog').addEventListener('close', () => {
    planoStopAlertCountdown_();
    planoLastAlertSendResult = null;
  });

  window.addEventListener('app-bootstrap-loaded', (event) => {
    planoApplyBootstrap(event.detail);
  });
  window.addEventListener('app-bootstrap-refreshing', () => {
    planoClearAccess();
  });
  window.addEventListener('app-bootstrap-failed', (event) => {
    planoApplyBootstrapFailure(event.detail);
  });

  if (window.__appBootstrapPayload) {
    planoApplyBootstrap(window.__appBootstrapPayload.plans);
  } else if (window.__appBootstrapError) {
    planoApplyBootstrapFailure(window.__appBootstrapError);
  }

  function planoShowView(view) {
    if (!window.__appAccessGranted) return;
    window.__activeAppView = view;
    const dashboardMain = document.getElementById('dashboard').closest('main');
    const plansView = document.getElementById('plans-view');
    const isPlans = view === 'plans';

    dashboardMain.classList.toggle('hidden', isPlans);
    plansView.classList.toggle('hidden', !isPlans);
    document.getElementById('nav-dashboard').classList.toggle('nav-button-active', !isPlans);
    document.getElementById('nav-plans').classList.toggle('nav-button-active', isPlans);
    document.getElementById('download-dashboard-button').classList.toggle('hidden', isPlans);
    if (isPlans) planoRenderAll();
  }

  async function planoLoadBootstrap() {
    try {
      const result = await planoCall('getPlanoAcaoBootstrap');
      planoApplyBootstrap(result);
    } catch (error) {
      planoApplyBootstrapFailure(planoErrorMessage(error));
    }
  }

  function planoApplyBootstrap(result) {
    const previousAccess = planoAccessSignature();
    planoState.available = Boolean(result && result.available);
    planoState.user = result && result.user ? result.user : null;
    planoState.permissions = result && result.permissions ? result.permissions : {};
    planoState.plans = result && Array.isArray(result.plans) ? result.plans : [];
    planoPruneDetailCache();
    planoState.options = result && result.options
      ? result.options
      : { statuses: [], priorities: [], responsibles: [] };
    planoState.message = result && result.message ? result.message : '';
    const currentAccess = planoAccessSignature();
    if (previousAccess && previousAccess !== currentAccess) planoResetSensitiveState();
    planoRenderAll();
  }

  function planoApplyBootstrapFailure(message) {
    planoClearAccess();
    planoState.message = String(message || 'Não foi possível validar o acesso ao sistema.');
    planoRenderAll();
  }

  function planoClearAccess() {
    planoResetSensitiveState();
    planoState.available = false;
    planoState.user = null;
    planoState.permissions = {};
    planoState.plans = [];
    planoState.options = { statuses: [], priorities: [], responsibles: [] };
  }

  function planoAccessSignature() {
    if (!planoState.user) return '';
    return JSON.stringify({
      email: planoState.user.email,
      role: planoState.user.role,
      regions: (planoState.user.regions || []).slice().sort(),
      permissions: planoState.permissions || {},
    });
  }

  function planoResetSensitiveState() {
    planoState.detailRequestId += 1;
    planoState.editingPlan = null;
    planoState.selectedDetail = null;
    planoState.selectedCity = '';
    planoState.attentionFilter = '';
    planoState.detailCache.clear();
    ['plan-form-dialog', 'plan-detail-dialog', 'plan-update-dialog', 'plan-image-dialog', 'alert-preview-dialog'].forEach((id) => {
      const dialog = document.getElementById(id);
      if (dialog && dialog.open) dialog.close();
    });
  }

  function planoRenderAll() {
    planoRenderAvailability();
    if (!planoState.available) return;
    planoRenderFilterOptions();
    planoRenderResponsibleOptions();
    planoRenderAttentionCenter();
    planoRenderSummary();
    planoRenderList();
    if (planoState.selectedCity) {
      const record = state.data && state.data.cities.find((item) => item.city === planoState.selectedCity);
      if (record) planoRenderCityActions(record);
    }
  }

  function planoRenderAvailability() {
    const unavailable = document.getElementById('plans-unavailable');
    const content = document.getElementById('plans-content');
    const summary = document.getElementById('plans-summary');
    const attention = document.getElementById('plans-attention');
    const newButton = document.getElementById('new-plan-button');
    const previewButton = document.getElementById('preview-alerts-button');
    const userLabel = document.getElementById('plans-user-label');

    if (!planoState.available) {
      unavailable.replaceChildren(
        planoElement('strong', '', 'Módulo ainda indisponível'),
        planoElement('span', '', planoState.message || 'Execute a função setupPlanoAcao_ no editor do Apps Script.')
      );
      unavailable.classList.remove('hidden');
      content.classList.add('hidden');
      summary.classList.add('hidden');
      attention.classList.add('hidden');
      newButton.classList.add('hidden');
      previewButton.classList.add('hidden');
      userLabel.textContent = 'Acesso ao sistema indisponível.';
      return;
    }

    unavailable.classList.add('hidden');
    content.classList.remove('hidden');
    summary.classList.remove('hidden');
    attention.classList.remove('hidden');
    newButton.classList.toggle('hidden', !planoState.permissions.create);
    previewButton.classList.toggle('hidden', !planoState.permissions.administer);
    userLabel.textContent = `${planoState.user.name} • ${planoRoleLabel(planoState.user.role)}`;
  }

  function planoRenderFilterOptions() {
    planoSetOptions(
      'plan-status-filter',
      [{ value: '', label: 'Todos os status' }].concat(planoState.options.statuses || [])
    );
    planoSetOptions(
      'plan-priority-filter',
      [{ value: '', label: 'Todas as prioridades' }].concat(planoState.options.priorities || [])
    );
    const regions = Array.from(new Set(planoState.plans.map((plan) => plan.regional).filter(Boolean)))
      .sort((a, b) => a.localeCompare(b, 'pt-BR'));
    planoSetOptions(
      'plan-region-filter',
      [{ value: '', label: 'Todas as regionais' }].concat(
        regions.map((regional) => ({ value: regional, label: regional }))
      )
    );
  }

  function planoRenderResponsibleOptions() {
    const datalist = document.getElementById('plan-responsible-options');
    const options = (planoState.options.responsibles || []).map((responsible) => {
      const option = document.createElement('option');
      option.value = responsible.email;
      option.label = `${responsible.name} — ${responsible.email}`;
      return option;
    });
    datalist.replaceChildren(...options);
  }

  function planoSetOptions(elementId, options, selectedValue) {
    const select = document.getElementById(elementId);
    const selected = selectedValue === undefined ? select.value : selectedValue;
    select.replaceChildren(...options.map((option) => createOption(option.value, option.label)));
    if (Array.from(select.options).some((option) => option.value === selected)) select.value = selected;
  }

  function planoRenderSummary() {
    const active = planoState.plans.filter(planoIsActive);
    const attention = planoState.plans.filter(planoNeedsAttention);
    const improvedCycles = Array.from(planoCreateCycleMap(planoState.plans).values())
      .filter((cycle) => cycle.state === 'final' && cycle.status === 'improved');
    const cards = [
      { label: 'Planos ativos', value: active.length },
      { label: 'Em atenção', value: attention.length, tone: attention.length ? 'negative' : 'positive' },
      { label: 'Ciclos com melhora', value: improvedCycles.length, tone: 'positive' },
      { label: 'Concluídos', value: planoState.plans.filter((plan) => plan.status === 'concluido').length, tone: 'positive' },
    ];
    document.getElementById('plans-summary').replaceChildren(...cards.map(createCard));
  }

  function planoRenderAttentionCenter() {
    const plans = planoState.plans;
    const options = [
      {
        value: 'attention',
        label: 'Em atenção',
        count: plans.filter(planoNeedsAttention).length,
        help: 'Planos únicos com ao menos um alerta',
        tone: 'danger',
      },
      {
        value: 'overdue',
        label: 'Atrasados',
        count: plans.filter(planoIsOverdue).length,
        help: 'Prazo vencido e plano ainda ativo',
        tone: 'danger',
      },
      {
        value: 'pending',
        label: 'Pendentes',
        count: plans.filter((plan) => planoAttentionFlags(plan).pending).length,
        help: 'Planos com status Pendente',
        tone: 'warning',
      },
      {
        value: 'stale',
        label: 'Sem atualização',
        count: plans.filter(planoIsStale).length,
        help: `${PLANO_STALE_DAYS} dias ou mais sem andamento`,
        tone: 'neutral',
      },
    ];
    const container = document.getElementById('plans-attention-options');
    container.replaceChildren(...options.map(planoCreateAttentionOption));
    document.getElementById('clear-attention-filter').disabled = !planoState.attentionFilter;
  }

  function planoCreateAttentionOption(option) {
    const button = planoElement('button', `attention-option attention-option-${option.tone}`);
    button.type = 'button';
    button.dataset.filter = option.value;
    button.setAttribute('aria-pressed', String(planoState.attentionFilter === option.value));
    if (planoState.attentionFilter === option.value) button.classList.add('attention-option-active');
    button.append(
      planoElement('strong', 'attention-option-count', String(option.count)),
      planoElement('span', 'attention-option-label', option.label),
      planoElement('small', 'attention-option-help', option.help)
    );
    button.addEventListener('click', () => planoSetAttentionFilter(option.value));
    return button;
  }

  function planoSetAttentionFilter(filter) {
    planoState.attentionFilter = planoState.attentionFilter === filter ? '' : filter;
    planoRenderAttentionCenter();
    planoRenderList();
  }

  function planoRenderList() {
    if (!planoState.available) return;
    const search = normalizeText(document.getElementById('plan-search').value);
    const status = document.getElementById('plan-status-filter').value;
    const regional = document.getElementById('plan-region-filter').value;
    const priority = document.getElementById('plan-priority-filter').value;
    const plans = planoState.plans.filter((plan) => {
      const haystack = normalizeText([plan.cidade, plan.o_que, plan.responsavel].join(' '));
      return (!search || haystack.includes(search))
        && (!status || plan.status === status)
        && (!regional || plan.regional === regional)
        && (!priority || plan.prioridade === priority)
        && planoMatchesAttentionFilter(plan, planoState.attentionFilter);
    });

    const rows = plans.length
      ? planoCreateGroupedListRows(plans)
      : [createEmptyRow(8, 'Nenhum plano encontrado.')];
    document.getElementById('plans-body').replaceChildren(...rows);
  }

  function planoCreateGroupedListRows(filteredPlans) {
    const allCityGroups = planoCreateCityMap(planoState.plans);
    const groups = new Map();
    filteredPlans.forEach((plan) => {
      const key = normalizeText(plan.cidade) || 'sem-cidade';
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(plan);
    });
    return Array.from(groups.entries()).flatMap(([key, visiblePlans]) => {
      const cityPlans = allCityGroups.get(key) || visiblePlans;
      const planRows = visiblePlans.map(planoCreateListRow);
      return [planoCreateCityGroupRow(cityPlans, visiblePlans.length), ...planRows];
    });
  }

  function planoCreateCityMap(plans) {
    const groups = new Map();
    (plans || []).forEach((plan) => {
      const key = normalizeText(plan.cidade) || 'sem-cidade';
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(plan);
    });
    return groups;
  }

  function planoCreateCityGroupRow(cityPlans, visiblePlanCount) {
    const row = document.createElement('tr');
    row.className = 'plan-cycle-row';
    const cell = appendCell(row, 'td', '');
    cell.colSpan = 8;
    const reference = cityPlans[0] || {};
    const activePlans = cityPlans.filter((plan) => plan.status !== 'cancelado');
    const progressPlans = activePlans.length ? activePlans : cityPlans;
    const averageProgress = progressPlans.length
      ? Math.round(progressPlans.reduce((sum, plan) => sum + (Number(plan.percentual_conclusao) || 0), 0) / progressPlans.length)
      : 0;
    const completedCount = cityPlans.filter((plan) => plan.status === 'concluido').length;
    const overdueCount = cityPlans.filter(planoIsOverdue).length;
    const periodCount = new Set(cityPlans.map((plan) => planoNormalizeMonthKey(plan.periodo_analisado)).filter(Boolean)).size;
    const summary = planoElement('div', 'plan-cycle-summary');
    const identity = planoElement('div', 'plan-cycle-identity');
    identity.append(
      planoElement('strong', 'plan-cycle-title', reference.cidade || 'Cidade não informada'),
      planoElement(
        'span',
        'plan-cycle-meta',
        visiblePlanCount === cityPlans.length
          ? `${cityPlans.length} planos em ${periodCount} ${periodCount === 1 ? 'período' : 'períodos'}`
          : `${visiblePlanCount} de ${cityPlans.length} planos exibidos`
      )
    );
    const metrics = planoElement('div', 'plan-cycle-metrics');
    metrics.append(
      planoElement('span', 'plan-cycle-chip', `Execução média ${averageProgress}%`),
      planoElement('span', 'plan-cycle-chip', `${completedCount} concluído${completedCount === 1 ? '' : 's'}`)
    );
    if (overdueCount) {
      metrics.appendChild(planoElement(
        'span',
        'plan-cycle-chip plan-cycle-chip-danger',
        `${overdueCount} atrasado${overdueCount === 1 ? '' : 's'}`
      ));
    }
    summary.append(identity, metrics);
    cell.appendChild(summary);
    return row;
  }

  function planoCreateListRow(plan) {
    const row = document.createElement('tr');
    row.tabIndex = 0;
    row.style.cursor = 'pointer';
    row.addEventListener('click', () => planoOpenDetail(plan.id));
    row.addEventListener('keydown', (event) => {
      if (event.key === 'Enter' || event.key === ' ') planoOpenDetail(plan.id);
    });

    const city = appendCell(row, 'td', '');
    city.append(
      planoElement('span', 'plan-row-city', plan.cidade),
      planoElement('small', 'plan-row-period', planoMonthLabel(plan.periodo_analisado))
    );
    appendCell(row, 'td', plan.o_que);
    appendCell(row, 'td', plan.responsavel);
    const priority = appendCell(row, 'td', planoPriorityLabel(plan.prioridade));
    priority.className = `plan-priority-${plan.prioridade}`;

    const status = appendCell(row, 'td', '');
    status.appendChild(planoStatusBadge(plan.status));
    const progress = appendCell(row, 'td', '');
    progress.appendChild(planoProgress(plan.percentual_conclusao));

    const deadline = appendCell(row, 'td', plan.prazo ? planoFormatDate(plan.prazo) : 'Sem prazo');
    if (planoIsOverdue(plan)) deadline.classList.add('plan-overdue');
    appendCell(row, 'td', planoFormatDateTime(plan.atualizado_em));
    applyMobileLabels(row, [
      'Cidade', 'Ação', 'Responsável', 'Prioridade', 'Status', 'Progresso', 'Prazo', 'Atualizado',
    ]);
    return row;
  }

  function planoRenderCityActions(record) {
    planoState.selectedCity = record.city;
    const container = document.getElementById('city-plan-actions');
    container.classList.remove('hidden');

    if (!planoState.available) {
      container.replaceChildren(planoElement('span', '', 'Planos de ação indisponíveis até executar setupPlanoAcao_.'));
      return;
    }

    const plans = planoState.plans.filter((plan) => plan.cidade === record.city);
    const active = plans.filter((plan) => !['concluido', 'cancelado'].includes(plan.status)).length;
    const summary = planoElement('div', 'city-plan-summary');
    summary.append(
      planoElement('strong', '', `${active} ${active === 1 ? 'plano ativo' : 'planos ativos'}`),
      planoElement('span', '', `${plans.length} no histórico desta cidade`)
    );
    const viewButton = planoElement('button', 'secondary-button', 'Ver planos');
    viewButton.type = 'button';
    viewButton.addEventListener('click', () => {
      closeCityDetail();
      planoShowView('plans');
      planoState.attentionFilter = '';
      planoRenderAttentionCenter();
      document.getElementById('plan-search').value = record.city;
      planoRenderList();
    });
    const children = [summary, viewButton];
    if (planoState.permissions.create) {
      const createButton = planoElement('button', 'primary-button', 'Criar plano');
      createButton.type = 'button';
      createButton.addEventListener('click', () => planoOpenForm(record.city));
      children.push(createButton);
    }
    container.replaceChildren(...children);
  }

  function planoOpenForm(city, plan) {
    if (!planoState.available || !planoState.permissions.create && !plan) return;
    planoState.editingPlan = plan || null;
    planoState.formOperationId = planoCreateOperationId();
    const form = document.getElementById('plan-form');
    form.reset();

    const cities = state.data
      ? state.data.cities.slice().sort((a, b) => a.city.localeCompare(b.city, 'pt-BR'))
      : [];
    if (plan && !cities.some((record) => record.city === plan.cidade)) {
      cities.push({ city: plan.cidade, regional: plan.regional });
    }
    planoSetOptions(
      'plan-city',
      cities.map((record) => ({ value: record.city, label: record.city })),
      plan ? plan.cidade : city || (cities[0] && cities[0].city)
    );
    document.getElementById('plan-city').disabled = Boolean(plan);
    planoSetOptions('plan-priority', planoState.options.priorities, plan ? plan.prioridade : 'media');
    planoSetOptions('plan-status', planoState.options.statuses, plan ? plan.status : 'nao_iniciado');

    if (plan) {
      Object.entries({
        produto: plan.produto,
        o_que: plan.o_que,
        como: plan.como,
        responsavel: plan.responsavel,
        responsavel_email: plan.responsavel_email || plan.criado_por,
        prazo: plan.prazo,
        percentual_conclusao: plan.percentual_conclusao,
        pendencia_motivo: plan.pendencia_motivo,
        proximo_passo: plan.proximo_passo,
        links_evidencias: (plan.links_evidencias || []).join('\n'),
      }).forEach(([name, value]) => {
        const field = form.elements.namedItem(name);
        if (field) field.value = value === undefined || value === null ? '' : value;
      });
    } else {
      form.elements.namedItem('produto').value = 'FTTH';
      form.elements.namedItem('percentual_conclusao').value = 0;
      if (planoState.user) form.elements.namedItem('responsavel').value = planoState.user.name;
      if (planoState.user) form.elements.namedItem('responsavel_email').value = planoState.user.email;
    }

    document.getElementById('plan-form-title').textContent = plan ? 'Editar plano de ação' : 'Novo plano de ação';
    const saveButton = document.getElementById('save-plan-button');
    saveButton.textContent = plan ? 'Salvar alterações' : 'Salvar plano';
    saveButton.dataset.label = saveButton.textContent;
    planoSyncPlanRegion();
    if (plan) document.getElementById('plan-regional').value = plan.regional || '';
    closeCityDetail();
    document.getElementById('plan-form-dialog').showModal();
  }

  function planoSyncResponsibleName() {
    const email = String(document.getElementById('plan-responsible-email').value || '').trim().toLowerCase();
    const responsible = (planoState.options.responsibles || []).find((item) => item.email === email);
    if (!responsible) return;
    document.getElementById('plan-form').elements.namedItem('responsavel').value = responsible.name;
  }

  async function planoPreviewAlerts() {
    const button = document.getElementById('preview-alerts-button');
    planoSetButtonBusy(button, true, 'Testando...');
    planoLastAlertSendResult = null;
    try {
      const result = await planoCall('previewPlanoAlerts');
      planoRenderAlertPreview(result);
      document.getElementById('alert-preview-dialog').showModal();
    } catch (error) {
      planoToast(planoErrorMessage(error));
    } finally {
      planoSetButtonBusy(button, false);
    }
  }

  async function planoSendAlertsNow() {
    const button = document.getElementById('send-alerts-now-button');
    // Reentrância/estado desatualizado: se o botão já está desabilitado (sem
    // pendente, ou um envio ainda em andamento), o clique não faz nada.
    if (button.disabled) return;
    if (!window.confirm('Isso envia e-mails reais agora, fora do horário agendado. Deseja continuar?')) return;
    if (!button.dataset.label) button.dataset.label = button.textContent;
    button.disabled = true;
    button.setAttribute('aria-busy', 'true');
    button.textContent = 'Enviando...';
    try {
      const result = await planoCall('executarAlertasAgora');
      planoLastAlertSendResult = result;
      planoToast(`Envio concluído: ${result.sentEmails || 0} e-mail(s), ${result.sentAlerts || 0} alerta(s).`);
      // Recarrega a simulação: ela é a única responsável por decidir se o
      // botão volta a ficar habilitado (só quando sobrar algo pendente de
      // verdade). Não usar planoSetButtonBusy(button, false) aqui: isso
      // reabilitaria o botão incondicionalmente e permitiria clicar de novo
      // com 0 alertas pendentes.
      const preview = await planoCall('previewPlanoAlerts');
      planoRenderAlertPreview(preview);
    } catch (error) {
      planoToast(planoErrorMessage(error));
      planoUpdateSendAlertsButton_(planoLastAlertPendingDeliveries);
    }
  }

  function planoRenderAlertPreview(result) {
    planoStopAlertCountdown_();
    const body = document.getElementById('alert-preview-body');
    const sendButton = document.getElementById('send-alerts-now-button');
    const canSendNow = Boolean(planoState.permissions && planoState.permissions.sendAlertsNow);
    sendButton.classList.toggle('hidden', !canSendNow);
    planoLastAlertPendingDeliveries = Number(result.pendingDeliveries) || 0;
    planoUpdateSendAlertsButton_(planoLastAlertPendingDeliveries);

    const fragments = [];
    if (planoLastAlertSendResult) fragments.push(planoAlertSendResultBanner(planoLastAlertSendResult));

    const summary = planoElement('section', 'alert-preview-summary');
    [
      { label: 'Planos ativos', value: result.activePlans || 0 },
      { label: 'Alertas detectados', value: result.detectedAlerts || 0, tone: result.detectedAlerts ? 'negative' : 'positive' },
      { label: 'Novos envios', value: result.pendingDeliveries || 0, tone: result.pendingDeliveries ? 'neutral' : 'positive' },
      { label: 'Sem destinatário', value: result.unresolvedAlerts || 0, tone: result.unresolvedAlerts ? 'negative' : 'positive' },
    ].forEach((card) => summary.appendChild(createCard(card)));
    fragments.push(summary);

    const note = planoElement('div', 'alert-preview-note');
    note.append(
      planoElement('strong', '', 'Simulação: nenhum e-mail foi enviado nesta consulta.'),
      planoElement(
        'span',
        'meta',
        `${result.suppressedDeliveries || 0} envio(s) já concluído(s) foram ignorados e ${result.recentAttempts || 0} tentativa(s) de hoje aguardam o próximo ciclo. Configuração: prazo em ${result.config.deadlineLeadDays} dia(s), sem atualização após ${result.config.staleDays} dias.`
      )
    );
    fragments.push(note);

    const list = planoElement('section', 'alert-preview-list');
    const items = result.items || [];
    if (!items.length) {
      list.appendChild(planoElement('p', 'alert-preview-empty', 'Nenhum novo alerta seria enviado agora.'));
    } else {
      items.forEach((item) => {
        const article = planoElement('article', 'alert-preview-item');
        const header = planoElement('div', 'alert-preview-item-header');
        header.append(
          planoElement('strong', '', item.city),
          planoElement('span', `plan-badge alert-type-${String(item.type || '').replace(/_/g, '-')}`, item.typeLabel)
        );
        article.append(
          header,
          planoElement('p', '', item.message),
          planoElement('span', 'meta', `Regional: ${item.regional || 'Não informada'}`),
          planoElement(
            'span',
            `meta ${item.recipients.length ? '' : 'negative'}`.trim(),
            item.recipients.length ? `Destinatários: ${item.recipients.join(', ')}` : 'Sem destinatário autorizado.'
          )
        );
        list.appendChild(article);
      });
    }
    fragments.push(list);

    body.replaceChildren(...fragments);
    planoStartAlertCountdown_(result);
  }

  function planoAlertSendResultBanner(result) {
    const hasIssue = (result.failedEmails || 0) > 0 || (result.unresolvedAlerts || 0) > 0;
    const banner = planoElement('div', `alert-preview-result${hasIssue ? ' negative' : ''}`);
    let detail = `${result.sentEmails || 0} e-mail(s) enviado(s), cobrindo ${result.sentAlerts || 0} alerta(s).`;
    if (result.failedEmails) detail += ` ${result.failedEmails} falha(s) de envio.`;
    if (result.unresolvedAlerts) detail += ` ${result.unresolvedAlerts} alerta(s) sem destinatário autorizado.`;
    banner.append(
      planoElement('strong', '', hasIssue ? 'Envio concluído com pendências.' : 'Envio concluído.'),
      planoElement('span', 'meta', detail)
    );
    return banner;
  }

  // Tempo restante até a próxima remessa automática (só mostra quando há
  // alertas pendentes; se não há nada pra enviar, o horário do gatilho é
  // irrelevante para o usuário). O alvo é um instante absoluto (ISO, calculado
  // no backend em America/Fortaleza), então a contagem no navegador funciona
  // corretamente em qualquer fuso do usuário.
  function planoStartAlertCountdown_(result) {
    const element = document.getElementById('alert-preview-countdown');
    const target = new Date(result.nextScheduledRunAt).getTime();
    if (!(result.pendingDeliveries > 0) || !Number.isFinite(target)) {
      element.classList.add('hidden');
      return;
    }
    const timeLabel = new Intl.DateTimeFormat('pt-BR', {
      timeZone: 'America/Fortaleza',
      hour: '2-digit',
      minute: '2-digit',
    }).format(new Date(target));
    const tick = () => {
      const diff = target - Date.now();
      element.textContent = diff <= 0
        ? `A remessa automática das ${timeLabel} pode iniciar a qualquer momento.`
        : `Próxima remessa automática em ${planoFormatCountdown(diff)} (às ${timeLabel}).`;
      element.classList.toggle('alert-preview-countdown-active', diff > 0);
    };
    tick();
    element.classList.remove('hidden');
    planoAlertCountdownTimer = window.setInterval(tick, 30000);
  }

  function planoUpdateSendAlertsButton_(pendingDeliveries) {
    // Único lugar que decide se "Enviar agora" pode ser clicado: só há algo
    // pendente de verdade. Chamado tanto pela renderização normal da
    // simulação quanto pelo tratamento de erro do próprio envio, para nunca
    // deixar o botão habilitado por engano.
    const button = document.getElementById('send-alerts-now-button');
    const hasPending = Number(pendingDeliveries) > 0;
    button.disabled = !hasPending;
    button.removeAttribute('aria-busy');
    if (button.dataset.label) button.textContent = button.dataset.label;
    button.title = hasPending ? '' : 'Nenhum alerta pendente para enviar agora.';
  }

  function planoStopAlertCountdown_() {
    if (!planoAlertCountdownTimer) return;
    window.clearInterval(planoAlertCountdownTimer);
    planoAlertCountdownTimer = null;
  }

  function planoFormatCountdown(ms) {
    const totalMinutes = Math.max(0, Math.round(ms / 60000));
    const hours = Math.floor(totalMinutes / 60);
    const minutes = totalMinutes % 60;
    if (hours <= 0) return `${minutes} min`;
    return minutes > 0 ? `${hours}h ${minutes}min` : `${hours}h`;
  }

  function planoSyncPlanRegion() {
    const city = document.getElementById('plan-city').value;
    const record = state.data && state.data.cities.find((item) => item.city === city);
    document.getElementById('plan-regional').value = record ? record.regional : '';
  }

  async function planoSubmitForm(event) {
    event.preventDefault();
    const form = event.currentTarget;
    const files = planoGetSelectedImages(form);
    try {
      planoValidateImages(files);
    } catch (error) {
      planoToast(planoErrorMessage(error));
      return;
    }
    const data = Object.fromEntries(new FormData(form).entries());
    delete data.imagens;
    const payload = Object.assign({}, data, {
      operacao_id: planoState.formOperationId,
      percentual_conclusao: Number(data.percentual_conclusao),
      links_evidencias: planoSplitLines(data.links_evidencias),
    });
    if (!planoState.editingPlan) {
      const record = state.data && state.data.cities.find((item) => item.city === data.cidade);
      if (!record) return planoToast('Cidade não encontrada no dashboard.');
      payload.cidade = record.city;
      payload.periodo_analisado = state.data.period.currentKey;
      payload.periodo_comparacao = state.data.period.previousKey;
    }

    const button = document.getElementById('save-plan-button');
    planoSetButtonBusy(button, true);
    try {
      let detail;
      if (planoState.editingPlan) {
        detail = await planoCall(
          'updatePlanoAcao',
          planoState.editingPlan.id,
          payload,
          planoState.editingPlan.versao
        );
      } else {
        detail = await planoCall('createPlanoAcao', payload);
      }
      const uploadResult = await planoUploadImages(detail.plan.id, '', files);
      document.getElementById('plan-form-dialog').close();
      const baseMessage = detail.noChanges
        ? 'Nenhuma alteração duplicada foi gravada.'
        : planoState.editingPlan ? 'Plano atualizado.' : 'Plano criado.';
      planoToast(planoBuildSaveMessage(baseMessage, uploadResult));
      planoState.editingPlan = null;
      planoState.formOperationId = '';
      await planoLoadBootstrap();
    } catch (error) {
      planoToast(planoErrorMessage(error));
    } finally {
      planoSetButtonBusy(button, false);
    }
  }

  async function planoOpenDetail(planId) {
    const dialog = document.getElementById('plan-detail-dialog');
    const summary = planoState.plans.find((plan) => String(plan.id) === String(planId));
    const cachedDetail = planoState.detailCache.get(String(planId));
    const requestId = ++planoState.detailRequestId;

    // Descarta observações de evidências do plano exibido anteriormente antes
    // de trocar o conteúdo do diálogo.
    if (planoEvidenceObserver) planoEvidenceObserver.disconnect();

    planoState.selectedDetail = cachedDetail || null;
    dialog.setAttribute('aria-busy', 'true');
    if (cachedDetail) {
      planoRenderDetail(cachedDetail);
    } else {
      planoRenderDetailLoading(summary);
    }
    if (!dialog.open) dialog.showModal();

    try {
      const detail = await planoCall('getPlanoAcaoDetail', planId);
      if (requestId !== planoState.detailRequestId) return;
      planoState.selectedDetail = detail;
      planoRenderDetail(detail);
    } catch (error) {
      if (requestId !== planoState.detailRequestId) return;
      planoRenderDetailError(summary, planId, error);
      planoToast(planoErrorMessage(error));
    } finally {
      if (requestId === planoState.detailRequestId) dialog.removeAttribute('aria-busy');
    }
  }

  function planoRenderDetailLoading(plan) {
    document.getElementById('plan-detail-title').textContent = plan ? plan.cidade : 'Plano de ação';
    document.getElementById('plan-detail-subtitle').textContent = 'Carregando detalhes...';

    const status = planoElement('div', 'plan-detail-loading');
    status.setAttribute('role', 'status');
    status.setAttribute('aria-live', 'polite');
    status.append(
      planoElement('span', 'plan-loading-spinner'),
      planoElement('strong', '', 'Carregando plano'),
      planoElement('span', 'meta', 'Buscando histórico, atualizações e evidências.')
    );

    const skeleton = planoElement('div', 'plan-detail-skeleton');
    for (let index = 0; index < 6; index += 1) {
      skeleton.appendChild(planoElement('span', 'plan-skeleton-line'));
    }
    document.getElementById('plan-detail-body').replaceChildren(status, skeleton);
  }

  function planoRenderDetailError(plan, planId, error) {
    document.getElementById('plan-detail-title').textContent = plan ? plan.cidade : 'Plano de ação';
    document.getElementById('plan-detail-subtitle').textContent = 'Não foi possível carregar os detalhes.';

    const container = planoElement('div', 'plan-detail-error');
    const retry = planoElement('button', 'primary-button', 'Tentar novamente');
    retry.type = 'button';
    retry.addEventListener('click', () => planoOpenDetail(planId));
    container.append(
      planoElement('strong', '', 'Falha ao carregar'),
      planoElement('span', 'meta', planoErrorMessage(error)),
      retry
    );
    document.getElementById('plan-detail-body').replaceChildren(container);
  }

  function planoRenderDetail(detail) {
    const plan = detail.plan;
    planoState.detailCache.set(String(plan.id), detail);
    document.getElementById('plan-detail-title').textContent = plan.cidade;
    document.getElementById('plan-detail-subtitle').textContent =
      `${plan.regional} • versão ${plan.versao} • atualizado ${planoFormatDateTime(plan.atualizado_em)}`;

    const body = document.getElementById('plan-detail-body');
    const metrics = planoElement('section', 'plan-detail-grid');
    [
      ['Status', planoStatusLabel(plan.status)],
      ['Prioridade', planoPriorityLabel(plan.prioridade)],
      ['Conclusão', `${plan.percentual_conclusao || 0}%`],
      ['Responsável', plan.responsavel],
      ['Prazo', plan.prazo ? planoFormatDate(plan.prazo) : 'Sem prazo'],
      ['Comparação', `${planoMonthLabel(plan.periodo_analisado)} x ${planoMonthLabel(plan.periodo_comparacao)}`],
    ].forEach(([label, value]) => {
      const card = planoElement('article', 'detail-card');
      card.append(planoElement('span', 'detail-card-label', label), planoElement('strong', 'detail-card-value', value || '—'));
      metrics.appendChild(card);
    });

    const actions = planoElement('div', 'plan-detail-actions');
    if (detail.permissions.update) {
      const edit = planoElement('button', 'secondary-button', 'Editar');
      edit.type = 'button';
      edit.addEventListener('click', () => {
        document.getElementById('plan-detail-dialog').close();
        planoOpenForm(plan.cidade, plan);
      });
      const update = planoElement('button', 'primary-button', 'Registrar andamento');
      update.type = 'button';
      update.addEventListener('click', () => planoOpenUpdate(detail));
      actions.append(edit, update);
    }
    if (detail.permissions.delete) {
      const remove = planoElement('button', 'danger-button', 'Excluir');
      remove.type = 'button';
      remove.addEventListener('click', () => planoDeletePlan(plan, remove));
      actions.prepend(remove);
    }

    const effectiveness = planoCycleEffectivenessBlock(plan);

    const what = planoDetailBlock('O que será feito', plan.o_que);
    const how = planoDetailBlock('Como será executado', plan.como);
    const next = planoDetailBlock('Próximo passo', plan.proximo_passo || 'Não informado');
    const pending = plan.pendencia_motivo
      ? planoDetailBlock('Motivo da pendência', plan.pendencia_motivo)
      : null;
    const evidence = (plan.links_evidencias || []).length
      ? planoLinkBlock('Evidências', plan.links_evidencias)
      : null;
    const imageEvidence = (detail.evidence || []).length
      ? planoEvidenceBlock(detail)
      : null;
    const timeline = planoElement('section', 'plan-detail-block');
    timeline.appendChild(planoElement('h3', '', 'Linha do tempo'));
    const timelineList = planoElement('div', 'plan-timeline');
    const updates = detail.updates || [];
    if (!updates.length) {
      timelineList.appendChild(planoElement('p', 'meta', 'Nenhuma atualização registrada.'));
    } else {
      updates.forEach((item) => {
        const entry = planoElement('article', 'plan-timeline-item');
        const links = planoTimelineLinks(item.links);
        entry.append(
          planoElement('strong', '', planoStatusLabel(item.status)),
          planoElement('p', '', item.resumo)
        );
        if (links) entry.appendChild(links);
        entry.appendChild(
          planoElement('span', 'plan-timeline-meta', `${item.percentual_conclusao}% • ${item.autor} • ${planoFormatDateTime(item.criado_em)}`)
        );
        timelineList.appendChild(entry);
      });
    }
    timeline.appendChild(timelineList);

    const audit = planoElement('section', 'plan-detail-block');
    audit.appendChild(planoElement('h3', '', 'Histórico de alterações'));
    const auditList = planoElement('div', 'plan-timeline');
    (detail.history || [])
      .filter((item) => item.operacao !== 'atualizacao')
      .slice(0, 20)
      .forEach((item) => {
      const entry = planoElement('article', 'plan-timeline-item');
      entry.append(
        planoElement('strong', '', planoOperationLabel(item.operacao)),
        planoElement('span', 'plan-timeline-meta', `${item.autor} • ${planoFormatDateTime(item.criado_em)}`)
      );
      auditList.appendChild(entry);
    });
    if (!auditList.children.length) {
      auditList.appendChild(planoElement('p', 'meta', 'Nenhuma alteração registrada.'));
    }
    audit.appendChild(auditList);

    body.replaceChildren(
      metrics,
      actions,
      effectiveness,
      what,
      how,
      next,
      ...(pending ? [pending] : []),
      ...(evidence ? [evidence] : []),
      ...(imageEvidence ? [imageEvidence] : []),
      timeline,
      audit
    );
  }

  function planoOpenUpdate(detail) {
    planoState.selectedDetail = detail;
    planoState.updateOperationId = planoCreateOperationId();
    const form = document.getElementById('plan-update-form');
    form.reset();
    planoSetOptions('plan-update-status', planoState.options.statuses, detail.plan.status);
    form.elements.namedItem('percentual_conclusao').value = detail.plan.percentual_conclusao || 0;
    form.elements.namedItem('pendencia_motivo').value = detail.plan.pendencia_motivo || '';
    form.elements.namedItem('proximo_passo').value = detail.plan.proximo_passo || '';
    document.getElementById('plan-detail-dialog').close();
    document.getElementById('plan-update-dialog').showModal();
  }

  async function planoSubmitUpdate(event) {
    event.preventDefault();
    const detail = planoState.selectedDetail;
    if (!detail) return;
    const files = planoGetSelectedImages(event.currentTarget);
    try {
      planoValidateImages(files);
    } catch (error) {
      planoToast(planoErrorMessage(error));
      return;
    }
    const data = Object.fromEntries(new FormData(event.currentTarget).entries());
    delete data.imagens;
    data.percentual_conclusao = Number(data.percentual_conclusao);
    data.links = planoSplitLines(data.links);
    data.operacao_id = planoState.updateOperationId;
    const button = document.getElementById('save-plan-update-button');
    planoSetButtonBusy(button, true);
    try {
      const updated = await planoCall('addPlanoAcaoUpdate', detail.plan.id, data, detail.plan.versao);
      const latestUpdate = updated.updates && updated.updates[0];
      const uploadResult = await planoUploadImages(
        updated.plan.id,
        updated.operationUpdateId || (latestUpdate ? latestUpdate.id : ''),
        files
      );
      const refreshed = uploadResult.successful
        ? await planoCall('getPlanoAcaoDetail', updated.plan.id)
        : updated;
      document.getElementById('plan-update-dialog').close();
      planoState.selectedDetail = refreshed;
      planoRenderDetail(refreshed);
      document.getElementById('plan-detail-dialog').showModal();
      const baseMessage = updated.noChanges
        ? 'Este andamento já havia sido registrado.'
        : 'Andamento registrado.';
      planoToast(planoBuildSaveMessage(baseMessage, uploadResult));
      planoState.updateOperationId = '';
      await planoLoadBootstrap();
    } catch (error) {
      planoToast(planoErrorMessage(error));
    } finally {
      planoSetButtonBusy(button, false);
    }
  }

  async function planoDeletePlan(plan, button) {
    const confirmed = window.confirm(
      `Excluir permanentemente o plano de ${plan.cidade}? Todos os registros e imagens vinculados serão removidos.`
    );
    if (!confirmed) return;
    planoSetButtonBusy(button, true, 'Excluindo...');
    try {
      await planoCall('deletePlanoAcao', plan.id, plan.versao);
      planoState.detailCache.delete(String(plan.id));
      document.getElementById('plan-detail-dialog').close();
      planoToast('Plano e registros vinculados excluídos.');
      await planoLoadBootstrap();
    } catch (error) {
      planoToast(planoErrorMessage(error));
    } finally {
      planoSetButtonBusy(button, false);
    }
  }

  function planoPruneDetailCache() {
    const currentVersions = new Map(
      planoState.plans.map((plan) => [String(plan.id), Number(plan.versao) || 1])
    );
    planoState.detailCache.forEach((detail, planId) => {
      const cachedVersion = detail && detail.plan ? Number(detail.plan.versao) || 1 : 0;
      if (!currentVersions.has(planId) || currentVersions.get(planId) !== cachedVersion) {
        planoState.detailCache.delete(planId);
      }
    });
  }

  function planoDetailBlock(title, content) {
    const block = planoElement('section', 'plan-detail-block');
    block.append(planoElement('h3', '', title), planoElement('p', '', content || '—'));
    return block;
  }

  function planoLinkBlock(title, links) {
    const block = planoElement('section', 'plan-detail-block');
    block.appendChild(planoElement('h3', '', title));
    const linkList = planoCreateLinkList(links, 'plan-link-list');
    if (linkList) block.appendChild(linkList);
    return block;
  }

  function planoTimelineLinks(links) {
    const linkList = planoCreateLinkList(links, 'plan-timeline-links');
    if (!linkList) return null;
    const wrapper = planoElement('div', 'plan-timeline-links-wrapper');
    wrapper.append(planoElement('span', 'plan-timeline-links-label', 'Links:'), linkList);
    return wrapper;
  }

  function planoCreateLinkList(links, className) {
    const normalizedLinks = Array.isArray(links) ? links.filter(Boolean) : [];
    if (!normalizedLinks.length) return null;
    const list = planoElement('div', className);
    normalizedLinks.forEach((url, index) => {
      const link = planoElement('a', 'plan-link', planoLinkLabel(url, index));
      link.href = url;
      link.target = '_blank';
      link.rel = 'noopener noreferrer';
      link.title = url;
      list.appendChild(link);
    });
    return list;
  }

  function planoLinkLabel(url, index) {
    try {
      return `${index + 1}. ${new URL(url).hostname}`;
    } catch (error) {
      return `Abrir link ${index + 1}`;
    }
  }

  // Cache em memória do conteúdo já baixado de cada evidência, por id, para
  // nunca buscar a mesma imagem duas vezes (nem ao reabrir o plano, nem ao
  // ampliar uma imagem já carregada pela miniatura).
  const planoEvidenceContentCache = new Map();

  // Só busca a imagem completa quando o card entra na área visível do diálogo,
  // em vez de disparar até 20 requisições em paralelo assim que o plano abre.
  const planoEvidenceObserver = 'IntersectionObserver' in window
    ? new IntersectionObserver((entries, observer) => {
      entries.forEach((entry) => {
        if (!entry.isIntersecting) return;
        observer.unobserve(entry.target);
        const { evidence, card, loading } = entry.target.__planoEvidence;
        planoFetchEvidenceImage(evidence, card, loading);
      });
    }, { root: null, rootMargin: '200px 0px', threshold: 0.01 })
    : null;

  function planoEvidenceBlock(detail) {
    const block = planoElement('section', 'plan-detail-block');
    block.appendChild(planoElement('h3', '', 'Imagens de evidência'));
    const gallery = planoElement('div', 'plan-evidence-gallery');
    (detail.evidence || []).forEach((evidence) => {
      const card = planoElement('article', 'plan-evidence-card');
      const loading = planoElement('div', 'plan-evidence-loading', 'Carregando imagem...');
      const caption = planoElement('div', 'plan-evidence-caption');
      caption.appendChild(planoElement('span', 'plan-evidence-name', evidence.nome));
      if (detail.permissions.update) {
        const remove = planoElement('button', 'plan-evidence-remove', 'Remover');
        remove.type = 'button';
        remove.addEventListener('click', () => planoRemoveEvidence(evidence.id, detail.plan.id));
        caption.appendChild(remove);
      }
      card.append(loading, caption);
      gallery.appendChild(card);
      planoLoadEvidenceImage(evidence, card, loading);
    });
    block.appendChild(gallery);
    return block;
  }

  function planoLoadEvidenceImage(evidence, card, loading) {
    const cached = planoEvidenceContentCache.get(evidence.id);
    if (cached) {
      planoRenderEvidenceImage(cached, card, loading);
      return;
    }
    if (!planoEvidenceObserver) {
      planoFetchEvidenceImage(evidence, card, loading);
      return;
    }
    card.__planoEvidence = { evidence, card, loading };
    planoEvidenceObserver.observe(card);
  }

  async function planoFetchEvidenceImage(evidence, card, loading) {
    try {
      const content = await planoCall('getPlanoEvidenceContent', evidence.id);
      planoEvidenceContentCache.set(evidence.id, content);
      planoRenderEvidenceImage(content, card, loading);
    } catch (error) {
      loading.textContent = 'Não foi possível carregar.';
      card.title = planoErrorMessage(error);
    }
  }

  function planoRenderEvidenceImage(content, card, loading) {
    const image = planoElement('img', 'plan-evidence-image');
    image.src = content.dataUrl;
    image.alt = content.name || 'Evidência do plano';
    image.loading = 'lazy';
    image.tabIndex = 0;
    image.setAttribute('role', 'button');
    image.setAttribute('aria-label', `Ampliar ${image.alt}`);
    loading.replaceWith(image);
    const openImage = () => planoOpenImage(content);
    image.addEventListener('click', openImage);
    image.addEventListener('keydown', (event) => {
      if (event.key !== 'Enter' && event.key !== ' ') return;
      event.preventDefault();
      openImage();
    });
    image.style.cursor = 'zoom-in';
  }

  function planoOpenImage(content) {
    const dialog = document.getElementById('plan-image-dialog');
    const image = document.getElementById('plan-image-preview');
    const name = content.name || 'Evidência';
    document.getElementById('plan-image-title').textContent = name;
    image.src = content.dataUrl;
    image.alt = name;
    if (!dialog.open) dialog.showModal();
  }

  async function planoRemoveEvidence(evidenceId, planId) {
    if (!window.confirm('Remover esta imagem?')) return;
    try {
      await planoCall('deletePlanoEvidence', evidenceId);
      planoToast('Imagem removida.');
      await planoOpenDetail(planId);
    } catch (error) {
      planoToast(planoErrorMessage(error));
    }
  }

  function planoGetSelectedImages(form) {
    const input = form.elements.namedItem('imagens');
    return input && input.files ? Array.from(input.files) : [];
  }

  function planoValidateImages(files) {
    if (files.length > PLANO_MAX_IMAGES_PER_UPLOAD) {
      throw new Error(`Selecione no máximo ${PLANO_MAX_IMAGES_PER_UPLOAD} imagens por envio.`);
    }
    files.forEach((file) => {
      if (!PLANO_ALLOWED_IMAGE_TYPES.includes(file.type)) {
        throw new Error(`Formato inválido em ${file.name}. Use JPG, PNG ou WEBP.`);
      }
      if (file.size > PLANO_MAX_IMAGE_BYTES) {
        throw new Error(`${file.name} ultrapassa o limite de 4 MB.`);
      }
    });
  }

  async function planoUploadImages(planId, updateId, files) {
    const result = { successful: 0, failed: [] };
    for (const file of files) {
      try {
        const payload = await planoReadImage(file);
        await planoCall('uploadPlanoEvidence', planId, updateId, payload);
        result.successful += 1;
      } catch (error) {
        result.failed.push({ name: file.name, message: planoErrorMessage(error) });
      }
    }
    return result;
  }

  function planoBuildSaveMessage(baseMessage, uploadResult) {
    const result = uploadResult || { successful: 0, failed: [] };
    const parts = [baseMessage];
    if (result.successful) parts.push(`${result.successful} imagem(ns) adicionada(s).`);
    if (result.failed.length) {
      parts.push(`${result.failed.length} imagem(ns) não foram enviadas. O plano foi salvo.`);
    }
    return parts.join(' ');
  }

  function planoCreateOperationId() {
    if (window.crypto && typeof window.crypto.randomUUID === 'function') {
      return window.crypto.randomUUID();
    }
    const random = window.crypto && typeof window.crypto.getRandomValues === 'function'
      ? window.crypto.getRandomValues(new Uint32Array(4))
      : [Math.random() * 0xffffffff, Math.random() * 0xffffffff, Date.now(), performance.now()];
    return `op_${Array.from(random, (value) => Math.floor(value).toString(36)).join('_')}`;
  }

  function planoReadImage(file) {
    return new Promise((resolve, reject) => {
      const reader = new FileReader();
      reader.addEventListener('load', () => {
        const dataUrl = String(reader.result || '');
        resolve({
          name: file.name,
          mimeType: file.type,
          base64: dataUrl.split(',')[1] || '',
        });
      });
      reader.addEventListener('error', () => reject(new Error(`Não foi possível ler ${file.name}.`)));
      reader.readAsDataURL(file);
    });
  }

  function planoStatusBadge(status) {
    return planoElement(
      'span',
      `plan-badge plan-status-${String(status || '').replace(/_/g, '-')}`,
      planoStatusLabel(status)
    );
  }

  function planoProgress(value) {
    const percentage = Math.max(0, Math.min(100, Number(value) || 0));
    const wrapper = planoElement('div', 'plan-progress');
    const track = planoElement('div', 'plan-progress-track');
    const bar = planoElement('div', 'plan-progress-bar');
    bar.style.width = `${percentage}%`;
    track.appendChild(bar);
    wrapper.append(track, planoElement('span', 'plan-progress-value', `${percentage}%`));
    return wrapper;
  }

  function planoStatusLabel(value) {
    const match = (planoState.options.statuses || []).find((option) => option.value === value);
    return match ? match.label : value || '—';
  }

  function planoPriorityLabel(value) {
    const match = (planoState.options.priorities || []).find((option) => option.value === value);
    return match ? match.label : value || '—';
  }

  function planoRoleLabel(role) {
    return { visualizador: 'Visualizador', editor: 'Editor', colaborador: 'Colaborador', gestor: 'Gestor', administrador: 'Administrador' }[role] || role;
  }

  function planoOperationLabel(operation) {
    return {
      criacao: 'Plano criado',
      edicao: 'Plano editado',
      atualizacao: 'Andamento registrado',
      exclusao: 'Plano excluído',
      evidencia_adicionada: 'Imagem adicionada',
      evidencia_removida: 'Imagem removida',
    }[operation] || operation || 'Alteração';
  }

  function planoIsOverdue(plan) {
    return Boolean(plan.prazo) && String(plan.prazo).slice(0, 10) < planoTodayKey() && planoIsActive(plan);
  }

  function planoIsActive(plan) {
    return !['concluido', 'cancelado'].includes(plan.status);
  }

  function planoIsStale(plan) {
    if (!planoIsActive(plan)) return false;
    const timestamp = plan.atualizado_em || plan.criado_em;
    if (!timestamp) return true;
    const updatedAt = new Date(timestamp);
    if (Number.isNaN(updatedAt.getTime())) return true;
    return Math.floor((Date.now() - updatedAt.getTime()) / 86400000) >= PLANO_STALE_DAYS;
  }

  function planoAttentionFlags(plan) {
    return {
      overdue: planoIsOverdue(plan),
      pending: planoIsActive(plan) && plan.status === 'pendente',
      stale: planoIsStale(plan),
    };
  }

  function planoNeedsAttention(plan) {
    const flags = planoAttentionFlags(plan);
    return flags.overdue || flags.pending || flags.stale;
  }

  function planoMatchesAttentionFilter(plan, filter) {
    if (!filter) return true;
    if (filter === 'attention') return planoNeedsAttention(plan);
    return Boolean(planoAttentionFlags(plan)[filter]);
  }

  function planoTodayParts_(now) {
    // Data "de hoje" no fuso America/Fortaleza (o mesmo do backend), nunca a
    // hora local do navegador. Base única usada tanto pelo marcador de dia
    // (planoTodayKey) quanto pela lógica de fechamento de ciclo
    // (planoCycleTiming), evitando duas fontes de "hoje" divergentes.
    const suppliedDate = now === undefined || now === null ? null : new Date(now);
    const source = suppliedDate && !Number.isNaN(suppliedDate.getTime()) ? suppliedDate : new Date();
    return new Intl.DateTimeFormat('en-CA', {
      timeZone: 'America/Fortaleza',
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
    }).formatToParts(source).reduce((parts, part) => {
      if (part.type !== 'literal') parts[part.type] = Number(part.value);
      return parts;
    }, {});
  }

  function planoTodayKey() {
    const { year, month, day } = planoTodayParts_();
    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }

  function planoCycleKey(plan) {
    const city = normalizeText(plan && plan.cidade) || 'sem-cidade';
    const monthKey = planoNormalizeMonthKey(plan && plan.periodo_analisado) || 'sem-periodo';
    return `${city}|${monthKey}`;
  }

  function planoCreateCycleMap(plans) {
    const groups = new Map();
    (plans || []).forEach((plan) => {
      const key = planoCycleKey(plan);
      if (!groups.has(key)) groups.set(key, []);
      groups.get(key).push(plan);
    });
    const cycles = new Map();
    groups.forEach((cyclePlans, key) => cycles.set(key, planoCreateCycle(cyclePlans)));
    return cycles;
  }

  function planoCreateCycle(plans) {
    const orderedPlans = [...(plans || [])].sort((a, b) => String(a.criado_em || '').localeCompare(String(b.criado_em || '')));
    const reference = orderedPlans[0] || {};
    const monthKey = planoNormalizeMonthKey(reference.periodo_analisado);
    const nextMonthKey = planoNextMonthKey(monthKey);
    const history = planoCityHistory(reference.cidade);
    const initialPoint = planoHistoryPoint(history, monthKey);
    const nextPoint = planoHistoryPoint(history, nextMonthKey);
    const planBaseline = orderedPlans
      .map((plan) => plan.base_atual)
      .find((value) => value !== '' && value !== null && value !== undefined && Number.isFinite(Number(value)));
    const baseline = planBaseline === undefined
      ? initialPoint ? Number(initialPoint.value) : Number.NaN
      : Number(planBaseline);
    const timing = planoCycleTiming(monthKey, Boolean(nextPoint));
    const observedPoint = timing.state === 'final' ? nextPoint : timing.state === 'preview' ? initialPoint : null;
    const current = Number(observedPoint && observedPoint.value);
    const available = Number.isFinite(baseline) && Number.isFinite(current) && Boolean(observedPoint);
    const difference = available ? current - baseline : null;
    const percentage = available && baseline ? difference / baseline : null;
    const status = !available ? 'unavailable' : difference > 0 ? 'improved' : difference < 0 ? 'worsened' : 'stable';
    const activeForProgress = orderedPlans.filter((plan) => plan.status !== 'cancelado');
    const progressPlans = activeForProgress.length ? activeForProgress : orderedPlans;
    const averageProgress = progressPlans.length
      ? Math.round(progressPlans.reduce((sum, plan) => sum + (Number(plan.percentual_conclusao) || 0), 0) / progressPlans.length)
      : 0;
    return {
      city: reference.cidade || 'Cidade não informada',
      regional: reference.regional || '',
      monthKey,
      nextMonthKey,
      state: timing.state,
      stateLabel: timing.label,
      availabilityDate: timing.availabilityDate,
      available,
      baseline,
      current,
      difference,
      percentage,
      status,
      label: status === 'improved' ? 'Melhorou' : status === 'worsened' ? 'Piorou' : status === 'stable' ? 'Sem mudança' : timing.label,
      observedLabel: observedPoint && (observedPoint.label || planoMonthLabel(observedPoint.monthKey)),
      planCount: orderedPlans.length,
      averageProgress,
      completedCount: orderedPlans.filter((plan) => plan.status === 'concluido').length,
      overdueCount: orderedPlans.filter(planoIsOverdue).length,
    };
  }

  function planoHistoryPoint(history, monthKey) {
    if (!history || !Array.isArray(history.points) || !monthKey) return null;
    return history.points.find((item) =>
      planoNormalizeMonthKey(item.monthKey) === monthKey
      && item.value !== ''
      && item.value !== null
      && item.value !== undefined
      && Number.isFinite(Number(item.value))
    ) || null;
  }

  function planoNextMonthKey(monthKey) {
    const match = String(monthKey || '').match(/^(\d{4})-(\d{2})$/);
    if (!match) return '';
    const date = new Date(Number(match[1]), Number(match[2]), 1);
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}`;
  }

  function planoCycleTiming(monthKey, hasNextMonth, now) {
    const match = String(monthKey || '').match(/^(\d{4})-(\d{2})$/);
    if (!match) return { state: 'waiting', label: 'Período não informado', availabilityDate: '' };

    const year = Number(match[1]);
    const month = Number(match[2]);
    const cycleIndex = year * 12 + month;
    const nextIndex = cycleIndex + 1;
    const today = planoTodayParts_(now);
    const currentIndex = today.year * 12 + today.month;

    // hasNextMonth só diz que já existe uma linha na base para o mês
    // seguinte — não que esse mês já fechou de verdade. Se a planilha for
    // alimentada de forma incremental, um valor parcial do mês em curso
    // poderia aparecer antes do fim do mês e virar "Resultado final" mostrando
    // um número que ainda vai mudar. Só confia nisso quando já estamos no mês
    // depois do mês seguinte (ou seja, o mês seguinte já terminou).
    if (hasNextMonth && currentIndex > nextIndex) {
      return { state: 'final', label: 'Resultado final', availabilityDate: '' };
    }

    if (cycleIndex < currentIndex) {
      return { state: 'waiting', label: `Aguardando base de ${planoMonthLabel(planoNextMonthKey(monthKey))}`, availabilityDate: '' };
    }

    const lastDay = new Date(year, month, 0).getDate();
    const previewDay = Math.max(1, lastDay - 1);
    const availabilityDate = `${year}-${String(month).padStart(2, '0')}-${String(previewDay).padStart(2, '0')}`;
    if (cycleIndex === currentIndex && today.day >= previewDay) {
      return { state: 'preview', label: 'Prévia de fechamento', availabilityDate };
    }
    return { state: 'monitoring', label: 'Em acompanhamento', availabilityDate };
  }

  function planoCityHistory(city) {
    if (!state.historyByCity) return null;
    const direct = state.historyByCity.get(city);
    if (direct) return direct;
    const normalizedCity = normalizeText(city);
    for (const [name, history] of state.historyByCity.entries()) {
      if (normalizeText(name) === normalizedCity) return history;
    }
    return null;
  }

  function planoNormalizeMonthKey(value) {
    const match = String(value || '').match(/^(\d{4})-(\d{2})/);
    return match ? `${match[1]}-${match[2]}` : '';
  }

  function planoCycleEffectivenessBlock(plan) {
    const effectiveness = planoCreateCycleMap(planoState.plans).get(planoCycleKey(plan)) || planoCreateCycle([plan]);
    const section = planoElement('section', 'plan-effectiveness-block');
    const header = planoElement('div', 'plan-effectiveness-header');
    header.append(
      planoElement('h3', '', 'Resultado consolidado do ciclo'),
      planoElement('p', 'meta', `${effectiveness.planCount} ${effectiveness.planCount === 1 ? 'plano criado' : 'planos criados'} para ${effectiveness.city} em ${planoMonthLabel(effectiveness.monthKey)}. A variação não comprova causalidade.`)
    );
    section.appendChild(header);
    const operationalCards = [
      ['Planos do ciclo', formatInteger(effectiveness.planCount), 'Agrupados por cidade e período'],
      ['Execução média', `${effectiveness.averageProgress}%`, 'Média dos planos não cancelados'],
      ['Concluídos', formatInteger(effectiveness.completedCount), 'Planos encerrados no ciclo'],
      ['Atrasados', formatInteger(effectiveness.overdueCount), 'Planos ativos com prazo vencido'],
    ];
    const grid = planoElement('div', 'plan-effectiveness-grid');
    operationalCards.forEach(([label, value, helper]) => grid.appendChild(planoEffectivenessCard(label, value, helper)));
    if (!effectiveness.available) {
      const message = effectiveness.state === 'monitoring'
        ? `Em acompanhamento. A prévia será liberada em ${planoFormatDate(effectiveness.availabilityDate)}, um dia antes do fechamento do mês.`
        : effectiveness.state === 'waiting'
          ? `${effectiveness.stateLabel}. O resultado final será calculado automaticamente assim que a nova base estiver disponível.`
          : effectiveness.stateLabel;
      section.append(grid, planoElement('p', 'plan-cycle-notice', message));
      return section;
    }
    const change = `${formatSignedInteger(effectiveness.difference)}${effectiveness.percentage === null ? '' : ` (${formatSignedPercent(effectiveness.percentage)})`}`;
    const cards = [
      ['Base inicial', formatInteger(effectiveness.baseline), planoMonthLabel(effectiveness.monthKey)],
      [effectiveness.state === 'preview' ? 'Base atual' : 'Base de fechamento', formatInteger(effectiveness.current), effectiveness.observedLabel],
      ['Resultado observado', change, effectiveness.state === 'preview' ? 'Valor provisório; pode mudar até o fechamento' : 'Base de fechamento − base inicial'],
      ['Situação', effectiveness.label, effectiveness.state === 'preview' ? 'Prévia do ciclo' : 'Resultado final do ciclo'],
    ];
    cards.forEach(([label, value, helper], index) => {
      grid.appendChild(planoEffectivenessCard(label, value, helper, index >= 2 ? effectiveness.status : ''));
    });
    section.appendChild(grid);
    return section;
  }

  function planoEffectivenessCard(label, value, helper, tone) {
    const card = planoElement('article', 'plan-effectiveness-card');
    const valueElement = planoElement('strong', 'plan-effectiveness-value', value);
    if (tone) valueElement.classList.add(`plan-effectiveness-tone-${tone}`);
    card.append(
      planoElement('span', 'plan-effectiveness-label', label),
      valueElement,
      planoElement('small', 'plan-effectiveness-help', helper)
    );
    return card;
  }

  function planoFormatDate(value) {
    if (!value) return '—';
    const parts = String(value).slice(0, 10).split('-');
    return parts.length === 3 ? `${parts[2]}/${parts[1]}/${parts[0]}` : value;
  }

  function planoFormatDateTime(value) {
    if (!value) return '—';
    const date = new Date(value);
    return Number.isNaN(date.getTime())
      ? String(value)
      : new Intl.DateTimeFormat('pt-BR', { dateStyle: 'short', timeStyle: 'short' }).format(date);
  }

  function planoMonthLabel(monthKey) {
    const normalizedMonth = String(monthKey || '').match(/^(\d{4})-(\d{2})/);
    const key = normalizedMonth ? `${normalizedMonth[1]}-${normalizedMonth[2]}` : '';
    const match = state.data && state.data.availableMonths.find((month) => month.key === key);
    if (match) return match.label;
    if (!key) return monthKey || '—';
    const months = [
      'Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
      'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro',
    ];
    const monthName = months[Number(normalizedMonth[2]) - 1];
    return monthName ? `${monthName} de ${normalizedMonth[1]}` : key;
  }

  function planoSplitLines(value) {
    return String(value || '').split(/\n+/).map((line) => line.trim()).filter(Boolean);
  }

  function planoElement(tagName, className, text) {
    const element = document.createElement(tagName);
    if (className) element.className = className;
    if (text !== undefined) element.textContent = text;
    return element;
  }

  function planoSetButtonBusy(button, busy, busyLabel) {
    button.disabled = busy;
    button.setAttribute('aria-busy', String(busy));
    if (!button.dataset.label) button.dataset.label = button.textContent;
    button.textContent = busy ? busyLabel || 'Salvando...' : button.dataset.label;
  }

  function planoToast(message) {
    const toast = document.getElementById('plans-toast');
    toast.textContent = message;
    toast.classList.remove('hidden');
    window.clearTimeout(planoToast.timeout);
    planoToast.timeout = window.setTimeout(() => toast.classList.add('hidden'), 5000);
  }

  function planoErrorMessage(error) {
    return error && error.message ? error.message : String(error || 'Erro inesperado.');
  }

  function planoCall(method, ...args) {
    return new Promise((resolve, reject) => {
      const runner = google.script.run
        .withSuccessHandler(resolve)
        .withFailureHandler(reject);
      switch (method) {
        case 'getPlanoAcaoBootstrap':
          runner.getPlanoAcaoBootstrap();
          break;
        case 'previewPlanoAlerts':
          runner.previewPlanoAlerts();
          break;
        case 'executarAlertasAgora':
          runner.executarAlertasAgora();
          break;
        case 'createPlanoAcao':
          runner.createPlanoAcao(args[0]);
          break;
        case 'updatePlanoAcao':
          runner.updatePlanoAcao(args[0], args[1], args[2]);
          break;
        case 'getPlanoAcaoDetail':
          runner.getPlanoAcaoDetail(args[0]);
          break;
        case 'addPlanoAcaoUpdate':
          runner.addPlanoAcaoUpdate(args[0], args[1], args[2]);
          break;
        case 'deletePlanoAcao':
          runner.deletePlanoAcao(args[0], args[1]);
          break;
        case 'getPlanoEvidenceContent':
          runner.getPlanoEvidenceContent(args[0]);
          break;
        case 'deletePlanoEvidence':
          runner.deletePlanoEvidence(args[0]);
          break;
        case 'uploadPlanoEvidence':
          runner.uploadPlanoEvidence(args[0], args[1], args[2]);
          break;
        default:
          reject(new Error(`Operação não permitida: ${method}`));
      }
    });
  }
</script>

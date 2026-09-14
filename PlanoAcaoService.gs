function getPlanoAcaoBootstrap() {
  try {
    return createPlanoAcaoBootstrap_(getAppUser_());
  } catch (error) {
    if (error instanceof PlanoUserError) {
      return { available: false, message: error.message };
    }
    // Erro não previsto (bug, falha de API do Google, etc.): registra o
    // detalhe técnico completo no backend e mostra apenas uma mensagem
    // genérica ao usuário.
    console.error(`Falha inesperada em getPlanoAcaoBootstrap: ${(error && error.stack) || error}`);
    return {
      available: false,
      message: 'Não foi possível carregar os planos de ação. Tente novamente em instantes.',
    };
  }
}

function createPlanoAcaoBootstrap_(user) {
  return {
    available: true,
    user: publicPlanoUser_(user),
    permissions: createPlanoPermissionFlags_(user),
    plans: listAccessiblePlans_(user).map(serializePlanoPublicPlan_),
    options: {
      statuses: PLANO_ACAO_CONFIG.STATUS.map((value) => ({ value, label: planoStatusLabel_(value) })),
      priorities: PLANO_ACAO_CONFIG.PRIORITIES.map((value) => ({ value, label: planoPriorityLabel_(value) })),
      responsibles: listPlanoResponsibleOptions_(user),
    },
  };
}

function getPlanoAcaoDetail(planId) {
  return planoRunPublic_('getPlanoAcaoDetail', () => getPlanoAcaoDetail_impl_(planId));
}

function getPlanoAcaoDetail_impl_(planId) {
  const user = getAppUser_();
  const plan = requireAccessiblePlan_(planId, user);
  const updates = readPlanoRecordsByField_(PLANO_ACAO_CONFIG.SHEETS.UPDATES, 'plano_id', plan.id)
    .sort(comparePlanoDatesDesc_)
    .map(serializePlanoPublicUpdate_);
  const history = readPlanoRecordsByField_(PLANO_ACAO_CONFIG.SHEETS.HISTORY, 'plano_id', plan.id)
    .sort(comparePlanoDatesDesc_)
    .filter((record) => String(record.operacao || '') !== 'atualizacao')
    .slice(0, 20)
    .map(serializePlanoPublicHistory_);
  const evidence = readPlanoRecordsByField_(PLANO_ACAO_CONFIG.SHEETS.EVIDENCE, 'plano_id', plan.id)
    .filter((record) => toPlanoBoolean_(record.ativo, true))
    .sort(comparePlanoDatesDesc_)
    .map(serializePlanoPublicEvidence_);

  return {
    plan: serializePlanoPublicPlan_(plan),
    updates,
    history,
    evidence,
    permissions: createPlanSpecificPermissionFlags_(user, plan),
  };
}

function serializePlanoPublicPlan_(record) {
  return serializePlanoPublicRecord_(record, PLANO_ACAO_CONFIG.PUBLIC_FIELDS.PLAN);
}

function serializePlanoPublicUpdate_(record) {
  return serializePlanoPublicRecord_(record, PLANO_ACAO_CONFIG.PUBLIC_FIELDS.UPDATE);
}

function serializePlanoPublicHistory_(record) {
  return serializePlanoPublicRecord_(record, PLANO_ACAO_CONFIG.PUBLIC_FIELDS.HISTORY);
}

function serializePlanoPublicEvidence_(record) {
  return serializePlanoPublicRecord_(record, PLANO_ACAO_CONFIG.PUBLIC_FIELDS.EVIDENCE);
}

function createPlanoAcao(payload) {
  return planoRunPublic_('createPlanoAcao', () => createPlanoAcao_impl_(payload));
}

function createPlanoAcao_impl_(payload) {
  const user = getAppUser_();
  requirePlanoPermission_(user, 'create');
  const operationId = validatePlanoOperationId_(payload && payload.operacao_id);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let planId = '';
  let duplicateRequest = false;
  try {
    ensurePlanoRuntimeSchema_();
    const duplicate = findPlanoRecordByOperation_(PLANO_ACAO_CONFIG.SHEETS.PLANS, operationId);
    if (duplicate) {
      requireOwnedPlanoOperation_(duplicate, user, operationId);
      planId = duplicate.id;
      duplicateRequest = true;
    } else {
      const normalized = validatePlanoPayload_(payload, null);
      Object.assign(normalized, createPlanoCreationSnapshot_(payload, user));
      const responsible = requirePlanoResponsible_(normalized.responsavel_email, normalized.regional);
      normalized.responsavel = responsible.name;
      normalized.responsavel_email = responsible.email;

      const now = new Date();
      const plan = Object.assign({}, normalized, {
        id: Utilities.getUuid(),
        operacao_id: operationId,
        links_evidencias_json: JSON.stringify(normalized.links_evidencias || []),
        contexto_criacao_json: JSON.stringify(normalized.contexto_criacao || {}),
        criado_por: user.email,
        criado_em: now,
        atualizado_por: user.email,
        atualizado_em: now,
        versao: 1,
        ativo: true,
      });
      delete plan.links_evidencias;
      delete plan.contexto_criacao;

      try {
        appendPlanoRecord_(getPlanoSheet_(PLANO_ACAO_CONFIG.SHEETS.PLANS), plan);
        addPlanoHistory_(
          plan.id,
          'criacao',
          null,
          serializePlanoRecord_(plan),
          '',
          user.email,
          operationId
        );
        SpreadsheetApp.flush();
        planId = plan.id;
      } catch (error) {
        rollbackPlanoCreate_(operationId);
        throw error;
      }
    }
  } finally {
    lock.releaseLock();
  }

  const detail = getPlanoAcaoDetail(planId);
  if (duplicateRequest) detail.noChanges = true;
  return detail;
}

function updatePlanoAcao(planId, payload, expectedVersion) {
  return planoRunPublic_('updatePlanoAcao', () => updatePlanoAcao_impl_(planId, payload, expectedVersion));
}

function updatePlanoAcao_impl_(planId, payload, expectedVersion) {
  const user = getAppUser_();
  requirePlanoPermission_(user, 'update');
  const operationId = validatePlanoOperationId_(payload && payload.operacao_id);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let duplicateRequest = false;

  try {
    ensurePlanoRuntimeSchema_();
    const duplicate = findPlanoRecordByOperation_(PLANO_ACAO_CONFIG.SHEETS.HISTORY, operationId);
    if (duplicate) {
      requireMatchingPlanoOperation_(duplicate, planId, operationId);
      requireAccessiblePlan_(planId, user);
      duplicateRequest = true;
    } else {
      const existing = requireAccessiblePlan_(planId, user);
      requirePlanWriteAccess_(user, existing);
      requirePlanoVersion_(existing, expectedVersion);

      const normalized = validatePlanoPayload_(payload, existing);
      const responsible = requirePlanoResponsible_(normalized.responsavel_email, existing.regional);
      normalized.responsavel = responsible.name;
      normalized.responsavel_email = responsible.email;
      if (!hasMeaningfulPlanoChanges_(existing, normalized)) {
        duplicateRequest = true;
      } else {
        const before = serializePlanoRecord_(existing);
        const now = new Date();
        const changes = Object.assign({}, normalized, {
          links_evidencias_json: JSON.stringify(normalized.links_evidencias || []),
          atualizado_por: user.email,
          atualizado_em: now,
          versao: (Number(existing.versao) || 1) + 1,
        });
        delete changes.links_evidencias;
        delete changes.contexto_criacao;
        removePlanoImmutableFields_(changes);

        try {
          updatePlanoRecord_(getPlanoSheet_(PLANO_ACAO_CONFIG.SHEETS.PLANS), existing.__row, changes);
          const after = Object.assign({}, before, serializePlanoRecord_(changes), { id: existing.id, ativo: true });
          addPlanoHistory_(
            existing.id,
            'edicao',
            before,
            after,
            planoLimitedText_(
              payload && payload.motivo,
              PLANO_ACAO_CONFIG.TEXT_LIMITS.historyReason,
              'Motivo da alteração'
            ),
            user.email,
            operationId
          );
          SpreadsheetApp.flush();
        } catch (error) {
          rollbackPlanoMutation_(existing, operationId);
          throw error;
        }
      }
    }
  } finally {
    lock.releaseLock();
  }

  const detail = getPlanoAcaoDetail(planId);
  if (duplicateRequest) detail.noChanges = true;
  return detail;
}

function addPlanoAcaoUpdate(planId, payload, expectedVersion) {
  return planoRunPublic_('addPlanoAcaoUpdate', () => addPlanoAcaoUpdate_impl_(planId, payload, expectedVersion));
}

function addPlanoAcaoUpdate_impl_(planId, payload, expectedVersion) {
  const user = getAppUser_();
  requirePlanoPermission_(user, 'update');
  const operationId = validatePlanoOperationId_(payload && payload.operacao_id);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let duplicateRequest = false;
  let operationUpdateId = '';

  try {
    ensurePlanoRuntimeSchema_();
    const duplicate = findPlanoRecordByOperation_(PLANO_ACAO_CONFIG.SHEETS.UPDATES, operationId);
    if (duplicate) {
      requireMatchingPlanoOperation_(duplicate, planId, operationId);
      requireAccessiblePlan_(planId, user);
      duplicateRequest = true;
      operationUpdateId = duplicate.id;
    } else {
      const existing = requireAccessiblePlan_(planId, user);
      requirePlanWriteAccess_(user, existing);
      requirePlanoVersion_(existing, expectedVersion);
      const update = validatePlanoUpdatePayload_(payload, existing);
      requireNonDuplicatePlanoUpdate_(existing.id, update, user.email);
      const before = serializePlanoRecord_(existing);
      const nextVersion = (Number(existing.versao) || 1) + 1;
      const now = new Date();

      const planChanges = {
        status: update.status,
        percentual_conclusao: update.percentual_conclusao,
        pendencia_motivo: update.pendencia_motivo,
        proximo_passo: update.proximo_passo,
        atualizado_por: user.email,
        atualizado_em: now,
        versao: nextVersion,
      };
      if (update.novo_prazo) planChanges.prazo = update.novo_prazo;

      const updateRecord = {
        id: Utilities.getUuid(),
        operacao_id: operationId,
        plano_id: existing.id,
        tipo: 'andamento',
        resumo: update.resumo,
        status: update.status,
        percentual_conclusao: update.percentual_conclusao,
        pendencia_motivo: update.pendencia_motivo,
        proximo_passo: update.proximo_passo,
        novo_prazo: update.novo_prazo || '',
        links_json: JSON.stringify(update.links || []),
        autor: user.email,
        criado_em: now,
        versao_plano_resultante: nextVersion,
      };
      operationUpdateId = updateRecord.id;
      try {
        updatePlanoRecord_(getPlanoSheet_(PLANO_ACAO_CONFIG.SHEETS.PLANS), existing.__row, planChanges);
        appendPlanoRecord_(getPlanoSheet_(PLANO_ACAO_CONFIG.SHEETS.UPDATES), updateRecord);
        const after = Object.assign({}, before, serializePlanoRecord_(planChanges));
        addPlanoHistory_(
          existing.id,
          'atualizacao',
          before,
          after,
          update.resumo,
          user.email,
          operationId
        );
        SpreadsheetApp.flush();
      } catch (error) {
        rollbackPlanoMutation_(existing, operationId);
        throw error;
      }
    }
  } finally {
    lock.releaseLock();
  }

  const detail = getPlanoAcaoDetail(planId);
  if (duplicateRequest) detail.noChanges = true;
  detail.operationUpdateId = operationUpdateId;
  return detail;
}

function uploadPlanoEvidence(planId, updateId, payload) {
  return planoRunPublic_('uploadPlanoEvidence', () => uploadPlanoEvidence_impl_(planId, updateId, payload));
}

function uploadPlanoEvidence_impl_(planId, updateId, payload) {
  const user = getAppUser_();
  requirePlanoPermission_(user, 'update');
  const fileData = validatePlanoEvidencePayload_(payload);
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  let driveFile = null;
  let evidence = null;
  let history = null;
  try {
    const plan = requireAccessiblePlan_(planId, user);
    requirePlanWriteAccess_(user, plan);
    requirePlanoUpdateBelongsToPlan_(updateId, plan.id);
    let bytes;
    try {
      bytes = Utilities.base64Decode(fileData.base64);
    } catch (decodeError) {
      throw new PlanoUserError('O conteúdo da imagem é inválido.');
    }
    if (bytes.length > PLANO_ACAO_CONFIG.MAX_EVIDENCE_BYTES) {
      throw new PlanoUserError('A imagem ultrapassa o limite de 4 MB.');
    }
    validatePlanoEvidenceSignature_(bytes, fileData.mimeType);
    const evidenceSheet = ensurePlanoSheet_(
      getSpreadsheet_(),
      PLANO_ACAO_CONFIG.SHEETS.EVIDENCE,
      PLANO_ACAO_CONFIG.HEADERS[PLANO_ACAO_CONFIG.SHEETS.EVIDENCE]
    );
    const checksum = Utilities.base64Encode(
      Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes)
    );
    const evidenceState = inspectPlanoEvidenceState_(plan.id, checksum);
    if (evidenceState.duplicate) return serializePlanoPublicEvidence_(evidenceState.duplicate);
    if (evidenceState.count >= PLANO_ACAO_CONFIG.MAX_EVIDENCE_COUNT_PER_PLAN) {
      throw new PlanoUserError(
        `O plano atingiu o limite de ${PLANO_ACAO_CONFIG.MAX_EVIDENCE_COUNT_PER_PLAN} imagens.`
      );
    }
    if (
      evidenceState.totalBytes + bytes.length
      > PLANO_ACAO_CONFIG.MAX_EVIDENCE_TOTAL_BYTES_PER_PLAN
    ) {
      throw new PlanoUserError('O plano atingiu o limite total de 40 MB em imagens.');
    }

    const blob = Utilities.newBlob(bytes, fileData.mimeType, fileData.name);
    try {
      driveFile = getPlanoEvidenceFolder_().createFile(blob);
    } catch (storageError) {
      console.error('Falha ao armazenar uma evidência no Drive.');
      throw new PlanoUserError('Não foi possível armazenar a imagem. Tente novamente.');
    }
    evidence = {
      id: Utilities.getUuid(),
      plano_id: plan.id,
      atualizacao_id: String(updateId || '').trim(),
      drive_file_id: driveFile.getId(),
      nome: fileData.name,
      tipo: fileData.mimeType,
      tamanho: bytes.length,
      checksum_sha256: checksum,
      url: '',
      autor: user.email,
      criado_em: new Date(),
      ativo: true,
    };
    evidence.__row = appendPlanoRecord_(evidenceSheet, evidence);
    history = addPlanoHistory_(plan.id, 'evidencia_adicionada', null, {
      id: evidence.id,
      nome: evidence.nome,
      tipo: evidence.tipo,
      tamanho: evidence.tamanho,
    }, '', user.email);
    SpreadsheetApp.flush();
    return serializePlanoPublicEvidence_(evidence);
  } catch (error) {
    safelyDeletePlanoRow_(PLANO_ACAO_CONFIG.SHEETS.HISTORY, history && history.__row);
    safelyDeletePlanoRow_(PLANO_ACAO_CONFIG.SHEETS.EVIDENCE, evidence && evidence.__row);
    if (driveFile) {
      try {
        driveFile.setTrashed(true);
      } catch (driveError) {
        console.error(`Falha ao remover evidência após erro: ${driveError.message}`);
      }
    }
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function getPlanoEvidenceContent(evidenceId) {
  return planoRunPublic_('getPlanoEvidenceContent', () => getPlanoEvidenceContent_impl_(evidenceId));
}

function getPlanoEvidenceContent_impl_(evidenceId) {
  const user = getAppUser_();
  const evidence = requirePlanoEvidence_(evidenceId);
  requireAccessiblePlan_(evidence.plano_id, user);
  let file;
  let blob;
  try {
    file = DriveApp.getFileById(String(evidence.drive_file_id));
    blob = file.getBlob();
  } catch (storageError) {
    console.error('Falha ao ler uma evidência no Drive.');
    throw new PlanoUserError('O arquivo desta evidência está indisponível.');
  }
  const bytes = blob.getBytes();
  const mimeType = String(evidence.tipo || blob.getContentType()).trim().toLowerCase();

  if (bytes.length > PLANO_ACAO_CONFIG.MAX_EVIDENCE_BYTES) {
    throw new PlanoUserError('A imagem excede o limite permitido para visualização.');
  }
  if (!PLANO_ACAO_CONFIG.ALLOWED_EVIDENCE_TYPES.includes(mimeType)) {
    throw new PlanoUserError('O formato desta evidência não é permitido para visualização.');
  }
  validatePlanoEvidenceSignature_(bytes, mimeType);

  return {
    id: evidence.id,
    name: String(evidence.nome || file.getName()),
    mimeType,
    dataUrl: `data:${mimeType};base64,${Utilities.base64Encode(bytes)}`,
  };
}

function deletePlanoEvidence(evidenceId) {
  return planoRunPublic_('deletePlanoEvidence', () => deletePlanoEvidence_impl_(evidenceId));
}

function deletePlanoEvidence_impl_(evidenceId) {
  const user = getAppUser_();
  requirePlanoPermission_(user, 'update');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const evidence = requirePlanoEvidence_(evidenceId);
    const plan = requireAccessiblePlan_(evidence.plano_id, user);
    requirePlanWriteAccess_(user, plan);
    let history = null;
    try {
      updatePlanoRecord_(getPlanoSheet_(PLANO_ACAO_CONFIG.SHEETS.EVIDENCE), evidence.__row, {
        ativo: false,
      });
      history = addPlanoHistory_(
        plan.id,
        'evidencia_removida',
        serializePlanoRecord_(evidence),
        null,
        '',
        user.email
      );
      SpreadsheetApp.flush();
    } catch (error) {
      safelyDeletePlanoRow_(PLANO_ACAO_CONFIG.SHEETS.HISTORY, history && history.__row);
      const original = Object.assign({}, evidence);
      delete original.__row;
      updatePlanoRecord_(getPlanoSheet_(PLANO_ACAO_CONFIG.SHEETS.EVIDENCE), evidence.__row, original);
      throw error;
    }

    try {
      DriveApp.getFileById(String(evidence.drive_file_id)).setTrashed(true);
    } catch (driveError) {
      // O registro lógico continua removido mesmo se o arquivo já não existir no Drive.
    }
    return { success: true, id: evidence.id };
  } finally {
    lock.releaseLock();
  }
}

function deletePlanoAcao(planId, expectedVersion) {
  return planoRunPublic_('deletePlanoAcao', () => deletePlanoAcao_impl_(planId, expectedVersion));
}

function deletePlanoAcao_impl_(planId, expectedVersion) {
  const user = getAppUser_();
  requirePlanoPermission_(user, 'delete');
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const existing = requireAccessiblePlan_(planId, user);
    requirePlanoVersion_(existing, expectedVersion);
    const backups = createPlanoDeletionBackup_(existing);
    let deleted;
    try {
      deleted = deletePlanoBackupRecords_(backups);
      SpreadsheetApp.flush();
    } catch (error) {
      restorePlanoDeletionBackup_(backups);
      throw error;
    }

    let trashedEvidenceFiles = 0;
    backups.evidence.forEach((item) => {
      const driveFileId = String(item.drive_file_id || '').trim();
      if (!driveFileId) return;
      try {
        DriveApp.getFileById(driveFileId).setTrashed(true);
        trashedEvidenceFiles += 1;
      } catch (driveError) {
        console.warn(`Não foi possível mover a evidência ${item.id || 'sem identificador'} para a lixeira.`);
      }
    });

    return {
      success: true,
      id: existing.id,
      deleted,
      trashedEvidenceFiles,
    };
  } finally {
    lock.releaseLock();
  }
}

function createPlanoDeletionBackup_(plan) {
  const spreadsheet = getSpreadsheet_();
  const alertSheetExists = Boolean(
    spreadsheet.getSheetByName(PLANO_ACAO_CONFIG.SHEETS.ALERT_HISTORY)
  );
  return {
    plan: [plan],
    updates: readPlanoRecordsByField_(PLANO_ACAO_CONFIG.SHEETS.UPDATES, 'plano_id', plan.id),
    history: readPlanoRecordsByField_(PLANO_ACAO_CONFIG.SHEETS.HISTORY, 'plano_id', plan.id),
    evidence: readPlanoRecordsByField_(PLANO_ACAO_CONFIG.SHEETS.EVIDENCE, 'plano_id', plan.id),
    alerts: alertSheetExists
      ? readPlanoRecordsByField_(PLANO_ACAO_CONFIG.SHEETS.ALERT_HISTORY, 'plano_id', plan.id)
      : [],
    alertSheetExists,
  };
}

function deletePlanoBackupRecords_(backup) {
  return {
    updates: deletePlanoRecordsByField_(PLANO_ACAO_CONFIG.SHEETS.UPDATES, 'plano_id', backup.plan[0].id),
    history: deletePlanoRecordsByField_(PLANO_ACAO_CONFIG.SHEETS.HISTORY, 'plano_id', backup.plan[0].id),
    evidence: deletePlanoRecordsByField_(PLANO_ACAO_CONFIG.SHEETS.EVIDENCE, 'plano_id', backup.plan[0].id),
    alerts: backup.alertSheetExists
      ? deletePlanoRecordsByField_(PLANO_ACAO_CONFIG.SHEETS.ALERT_HISTORY, 'plano_id', backup.plan[0].id)
      : 0,
    plans: deletePlanoRows_(getPlanoSheet_(PLANO_ACAO_CONFIG.SHEETS.PLANS), [backup.plan[0].__row]),
  };
}

function restorePlanoDeletionBackup_(backup) {
  const planId = backup.plan[0].id;
  const groups = [
    [PLANO_ACAO_CONFIG.SHEETS.PLANS, 'id', backup.plan],
    [PLANO_ACAO_CONFIG.SHEETS.UPDATES, 'plano_id', backup.updates],
    [PLANO_ACAO_CONFIG.SHEETS.HISTORY, 'plano_id', backup.history],
    [PLANO_ACAO_CONFIG.SHEETS.EVIDENCE, 'plano_id', backup.evidence],
  ];
  if (backup.alertSheetExists) {
    groups.push([PLANO_ACAO_CONFIG.SHEETS.ALERT_HISTORY, 'plano_id', backup.alerts]);
  }

  groups.forEach(([sheetName, fieldName, records]) => {
    try {
      deletePlanoRecordsByField_(sheetName, fieldName, planId);
      const sheet = getPlanoSheet_(sheetName);
      records.forEach((record) => appendPlanoRecord_(sheet, record));
    } catch (rollbackError) {
      console.error(`Falha ao restaurar registros em ${sheetName}: ${rollbackError.message}`);
    }
  });
  SpreadsheetApp.flush();
}

function getAppUser_() {
  const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  if (!email) {
    throw new PlanoUserError(
      'Usuário não identificado. Publique o app somente para o domínio e valide a identificação do Google Workspace.'
    );
  }

  const matchingRecords = readPlanoRecords_(PLANO_ACAO_CONFIG.SHEETS.USERS).filter((item) =>
    String(item.email || '').trim().toLowerCase() === email && toPlanoBoolean_(item.ativo, false)
  );
  if (!matchingRecords.length) throw new PlanoUserError('Usuário sem acesso ao sistema.');
  if (matchingRecords.length > 1) {
    throw new PlanoUserError('Cadastro de usuário duplicado. Corrija a aba Usuarios_Permissoes.');
  }

  const record = matchingRecords[0];
  const role = normalizeAppRole_(record.papel);
  if (!PLANO_ACAO_CONFIG.ROLES.includes(role)) throw new PlanoUserError(`Papel inválido para ${email}: ${role}`);
  const regions = Array.from(new Set(
    parsePlanoJson_(record.regionais_json, []).map(normalizeAppRegional_).filter(Boolean)
  ));
  if (role !== 'administrador' && !regions.length) {
    throw new PlanoUserError('Usuário sem regionais configuradas.');
  }

  return {
    email,
    name: String(record.nome || email).trim(),
    role,
    regions,
  };
}

function normalizeAppRole_(role) {
  const normalizedRole = String(role || '').trim().toLowerCase();
  return ['colaborador', 'gestor'].includes(normalizedRole) ? 'editor' : normalizedRole;
}

function normalizeAppRegional_(regional) {
  return String(regional || '').trim().toUpperCase();
}

function publicPlanoUser_(user) {
  return { email: user.email, name: user.name, role: user.role, regions: user.regions };
}

function createPlanoPermissionFlags_(user) {
  return {
    view: true,
    create: user.role !== 'visualizador',
    update: ['editor', 'administrador'].includes(user.role),
    delete: user.role === 'administrador',
    administer: user.role === 'administrador',
    // Disparo manual de e-mails reais (fora do horário agendado) fica restrito
    // ao dono técnico do deploy, o mesmo guard que executarAlertasAgora já
    // aplica no backend — este flag só existe para a UI decidir se mostra o
    // botão, a permissão de fato é sempre revalidada no servidor.
    sendAlertsNow: user.role === 'administrador' && isDeploymentOwner_(),
  };
}

function listPlanoResponsibleOptions_(currentUser) {
  return readPlanoRecords_(PLANO_ACAO_CONFIG.SHEETS.USERS)
    .filter((record) => toPlanoBoolean_(record.ativo, false))
    .map(createPlanoUserFromRecord_)
    .filter((candidate) => candidate && (
      currentUser.role === 'administrador'
      || candidate.role === 'administrador'
      || candidate.regions.includes('*')
      || candidate.regions.some((regional) => currentUser.regions.includes(regional))
    ))
    .map((candidate) => ({ email: candidate.email, name: candidate.name }))
    .sort((a, b) => a.name.localeCompare(b.name, 'pt-BR'));
}

function requirePlanoResponsible_(email, regional) {
  const normalizedEmail = String(email || '').trim().toLowerCase();
  if (!/^\S+@\S+\.\S+$/.test(normalizedEmail)) {
    throw new PlanoUserError('Informe um e-mail válido para o responsável.');
  }
  const matches = readPlanoRecords_(PLANO_ACAO_CONFIG.SHEETS.USERS)
    .filter((record) => toPlanoBoolean_(record.ativo, false))
    .filter((record) => String(record.email || '').trim().toLowerCase() === normalizedEmail);
  if (matches.length !== 1) {
    throw new PlanoUserError('O e-mail do responsável deve pertencer a um único usuário ativo em Usuarios_Permissoes.');
  }
  const responsible = createPlanoUserFromRecord_(matches[0]);
  if (!responsible || !hasAppRegionalAccess_(responsible, regional)) {
    throw new PlanoUserError(`O responsável não possui acesso à regional ${regional || 'informada'}.`);
  }
  return responsible;
}

function createPlanoUserFromRecord_(record) {
  const email = String(record && record.email || '').trim().toLowerCase();
  const role = normalizeAppRole_(record && record.papel);
  const rawRole = String(record && record.papel || '').trim().toLowerCase();
  if (!email || !PLANO_ACAO_CONFIG.ROLES.includes(role)) return null;
  return {
    email,
    name: String(record.nome || email).trim(),
    role,
    rawRole,
    receivesManagementAlerts: toPlanoBoolean_(
      record.recebe_alertas_gestao,
      rawRole === 'gestor' || role === 'administrador'
    ),
    regions: Array.from(new Set(
      parsePlanoJson_(record.regionais_json, []).map(normalizeAppRegional_).filter(Boolean)
    )),
  };
}

function createPlanSpecificPermissionFlags_(user, plan) {
  const permissions = createPlanoPermissionFlags_(user);
  const canWrite = permissions.update && hasAppRegionalAccess_(user, plan.regional);
  return {
    view: true,
    update: canWrite,
    addUpdate: canWrite,
    delete: permissions.delete,
  };
}

function requirePlanoPermission_(user, permission) {
  if (!createPlanoPermissionFlags_(user)[permission]) {
    throw new PlanoUserError(`Você não possui permissão para ${permission === 'delete' ? 'excluir' : 'alterar'} planos de ação.`);
  }
}

function listAccessiblePlans_(user) {
  return readPlanoRecords_(PLANO_ACAO_CONFIG.SHEETS.PLANS)
    .filter((plan) => toPlanoBoolean_(plan.ativo, true) && hasAppRegionalAccess_(user, plan.regional))
    .sort(comparePlanoDatesDesc_);
}

function requireAccessiblePlan_(planId, user) {
  const plan = findPlanoRecordById_(PLANO_ACAO_CONFIG.SHEETS.PLANS, planId);
  if (!plan || !toPlanoBoolean_(plan.ativo, true) || !hasAppRegionalAccess_(user, plan.regional)) {
    throw new PlanoUserError('Plano de ação não encontrado ou sem permissão de acesso.');
  }
  return plan;
}

function requirePlanWriteAccess_(user, plan) {
  requireAppRegionalAccess_(user, plan.regional);
}

function hasAppRegionalAccess_(user, regional) {
  const normalizedRegional = normalizeAppRegional_(regional);
  return user.role === 'administrador'
    || user.regions.includes('*')
    || user.regions.includes(normalizedRegional);
}

function requireAppRegionalAccess_(user, regional) {
  if (!hasAppRegionalAccess_(user, regional)) {
    throw new PlanoUserError(`Você não possui permissão para a regional ${regional || 'não informada'}.`);
  }
}

function requirePlanoVersion_(plan, expectedVersion) {
  if ((Number(plan.versao) || 1) !== (Number(expectedVersion) || 0)) {
    throw new PlanoUserError('Este plano foi alterado por outra pessoa. Reabra-o antes de salvar novamente.');
  }
}

function validatePlanoPayload_(payload, existing) {
  const source = existing
    ? Object.assign({}, payload || {}, {
      cidade: existing.cidade,
      regional: existing.regional,
      periodo_analisado: existing.periodo_analisado,
      periodo_comparacao: existing.periodo_comparacao,
      base_anterior: existing.base_anterior,
      base_atual: existing.base_atual,
      diferenca: existing.diferenca,
      variacao_pct: existing.variacao_pct,
      contexto_criacao: parsePlanoJson_(existing.contexto_criacao_json, {}),
    })
    : Object.assign({}, payload || {});
  const percentage = clampPlanoPercentage_(source.percentual_conclusao);
  const status = validatePlanoOption_(source.status || 'nao_iniciado', PLANO_ACAO_CONFIG.STATUS, 'status');
  const normalized = {
    cidade: normalizeCity_(source.cidade),
    regional: String(source.regional || '').trim(),
    produto: planoLimitedText_(
      source.produto || 'FTTH',
      PLANO_ACAO_CONFIG.TEXT_LIMITS.product,
      'Produto'
    ),
    periodo_analisado: normalizeMonthKey_(source.periodo_analisado),
    periodo_comparacao: normalizeMonthKey_(source.periodo_comparacao),
    base_anterior: planoNumber_(source.base_anterior),
    base_atual: planoNumber_(source.base_atual),
    diferenca: planoNumber_(source.diferenca),
    variacao_pct: planoNumber_(source.variacao_pct),
    prioridade: validatePlanoOption_(source.prioridade || 'media', PLANO_ACAO_CONFIG.PRIORITIES, 'prioridade'),
    o_que: planoLimitedText_(source.o_que, PLANO_ACAO_CONFIG.TEXT_LIMITS.action, 'O que será feito'),
    como: planoLimitedText_(source.como, PLANO_ACAO_CONFIG.TEXT_LIMITS.execution, 'Como será executado'),
    responsavel: planoLimitedText_(
      source.responsavel,
      PLANO_ACAO_CONFIG.TEXT_LIMITS.responsibleName,
      'Responsável'
    ),
    responsavel_email: planoLimitedText_(
      source.responsavel_email || source.criado_por,
      PLANO_ACAO_CONFIG.TEXT_LIMITS.email,
      'E-mail do responsável'
    ).toLowerCase(),
    prazo: validatePlanoDate_(source.prazo),
    status,
    percentual_conclusao: status === 'concluido' ? 100 : percentage,
    pendencia_motivo: planoLimitedText_(
      source.pendencia_motivo,
      PLANO_ACAO_CONFIG.TEXT_LIMITS.pendingReason,
      'Motivo da pendência'
    ),
    proximo_passo: planoLimitedText_(
      source.proximo_passo,
      PLANO_ACAO_CONFIG.TEXT_LIMITS.nextStep,
      'Próximo passo'
    ),
    links_evidencias: normalizePlanoLinks_(source.links_evidencias || source.links_evidencias_json),
    contexto_criacao: source.contexto_criacao || parsePlanoJson_(source.contexto_criacao_json, {}),
  };

  if (!normalized.cidade) throw new PlanoUserError('Informe a cidade.');
  if (!normalized.o_que) throw new PlanoUserError('Informe o que será feito.');
  if (!normalized.como) throw new PlanoUserError('Informe como a ação será executada.');
  if (!normalized.responsavel_email) throw new PlanoUserError('Informe o e-mail do responsável.');
  if (normalized.status === 'pendente' && !normalized.pendencia_motivo) {
    throw new PlanoUserError('Informe o motivo da pendência.');
  }
  return normalized;
}

function validatePlanoUpdatePayload_(payload, plan) {
  const source = payload || {};
  const status = validatePlanoOption_(source.status || plan.status, PLANO_ACAO_CONFIG.STATUS, 'status');
  const update = {
    resumo: planoLimitedText_(
      source.resumo,
      PLANO_ACAO_CONFIG.TEXT_LIMITS.updateSummary,
      'Resumo da atualização'
    ),
    status,
    percentual_conclusao: status === 'concluido' ? 100 : clampPlanoPercentage_(source.percentual_conclusao),
    pendencia_motivo: planoLimitedText_(
      source.pendencia_motivo,
      PLANO_ACAO_CONFIG.TEXT_LIMITS.pendingReason,
      'Motivo da pendência'
    ),
    proximo_passo: planoLimitedText_(
      source.proximo_passo,
      PLANO_ACAO_CONFIG.TEXT_LIMITS.nextStep,
      'Próximo passo'
    ),
    novo_prazo: validatePlanoDate_(source.novo_prazo),
    links: normalizePlanoLinks_(source.links),
  };
  if (!update.resumo) throw new PlanoUserError('Descreva a atualização realizada.');
  if (update.status === 'pendente' && !update.pendencia_motivo) {
    throw new PlanoUserError('Informe o motivo da pendência.');
  }
  return update;
}

function hasMeaningfulPlanoChanges_(existing, normalized) {
  const existingLinks = parsePlanoJson_(existing.links_evidencias_json, []);
  const fields = [
    'produto', 'prioridade', 'o_que', 'como', 'responsavel', 'responsavel_email',
    'prazo', 'status', 'percentual_conclusao',
    'pendencia_motivo', 'proximo_passo',
  ];
  const changedField = fields.some((field) =>
    canonicalPlanoValue_(existing[field]) !== canonicalPlanoValue_(normalized[field])
  );
  return changedField
    || JSON.stringify(existingLinks) !== JSON.stringify(normalized.links_evidencias || []);
}

function removePlanoImmutableFields_(record) {
  [
    'cidade',
    'regional',
    'periodo_analisado',
    'periodo_comparacao',
    'base_anterior',
    'base_atual',
    'diferenca',
    'variacao_pct',
    'contexto_criacao_json',
  ].forEach((field) => delete record[field]);
}

function requireNonDuplicatePlanoUpdate_(planId, update, email) {
  const latest = readPlanoRecordsByField_(PLANO_ACAO_CONFIG.SHEETS.UPDATES, 'plano_id', planId)
    .sort(comparePlanoDatesDesc_)[0];
  if (!latest || String(latest.autor || '').toLowerCase() !== email) return;

  const createdAt = new Date(latest.criado_em).getTime();
  if (!createdAt || Date.now() - createdAt > 10 * 60 * 1000) return;
  const same = canonicalPlanoValue_(latest.resumo) === canonicalPlanoValue_(update.resumo)
    && canonicalPlanoValue_(latest.status) === canonicalPlanoValue_(update.status)
    && Number(latest.percentual_conclusao) === Number(update.percentual_conclusao)
    && canonicalPlanoValue_(latest.pendencia_motivo) === canonicalPlanoValue_(update.pendencia_motivo)
    && canonicalPlanoValue_(latest.proximo_passo) === canonicalPlanoValue_(update.proximo_passo)
    && canonicalPlanoValue_(latest.novo_prazo) === canonicalPlanoValue_(update.novo_prazo)
    && JSON.stringify(parsePlanoJson_(latest.links_json, [])) === JSON.stringify(update.links || []);
  if (same) throw new PlanoUserError('Esta atualização já foi registrada.');
}

function canonicalPlanoValue_(value) {
  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    return Utilities.formatDate(value, getPlanoTimeZone_(), 'yyyy-MM-dd');
  }
  if (typeof value === 'number') return String(value);
  return String(value === null || value === undefined ? '' : value).trim();
}

function validatePlanoEvidencePayload_(payload) {
  const source = payload || {};
  const mimeType = String(source.mimeType || '').trim().toLowerCase();
  if (!PLANO_ACAO_CONFIG.ALLOWED_EVIDENCE_TYPES.includes(mimeType)) {
    throw new PlanoUserError('Formato inválido. Use JPG, PNG ou WEBP.');
  }

  const name = String(source.name || 'evidencia')
    .replace(/[\\/:*?"<>|\u0000-\u001f]/g, '_')
    .trim()
    .slice(0, 150);
  const base64 = String(source.base64 || '').replace(/^data:[^;]+;base64,/, '').trim();
  if (!base64) throw new PlanoUserError('A imagem está vazia.');
  if (base64.length > Math.ceil(PLANO_ACAO_CONFIG.MAX_EVIDENCE_BYTES * 4 / 3) + 16) {
    throw new PlanoUserError('A imagem ultrapassa o limite de 4 MB.');
  }
  return { name: name || 'evidencia', mimeType, base64 };
}

function requirePlanoEvidence_(evidenceId) {
  const evidence = findPlanoRecordById_(PLANO_ACAO_CONFIG.SHEETS.EVIDENCE, evidenceId);
  if (!evidence || !toPlanoBoolean_(evidence.ativo, true)) {
    throw new PlanoUserError('Evidência não encontrada.');
  }
  return evidence;
}

function inspectPlanoEvidenceState_(planId, checksum) {
  const activeEvidence = readPlanoRecordsByField_(
    PLANO_ACAO_CONFIG.SHEETS.EVIDENCE,
    'plano_id',
    planId
  ).filter((record) => toPlanoBoolean_(record.ativo, true));
  return {
    count: activeEvidence.length,
    totalBytes: activeEvidence.reduce((total, record) => total + (Number(record.tamanho) || 0), 0),
    duplicate: activeEvidence.find((record) => String(record.checksum_sha256 || '') === checksum) || null,
  };
}

function validatePlanoEvidenceSignature_(bytes, mimeType) {
  const source = bytes || [];
  const data = [];
  for (let index = 0; index < Math.min(source.length, 12); index += 1) {
    data.push(Number(source[index]) & 0xff);
  }
  const matches = mimeType === 'image/jpeg'
    ? data.length >= 3 && data[0] === 0xff && data[1] === 0xd8 && data[2] === 0xff
    : mimeType === 'image/png'
      ? [0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]
        .every((value, index) => data[index] === value)
      : mimeType === 'image/webp'
        ? data.length >= 12
          && String.fromCharCode(...data.slice(0, 4)) === 'RIFF'
          && String.fromCharCode(...data.slice(8, 12)) === 'WEBP'
        : false;
  if (!matches) {
    throw new PlanoUserError('O conteúdo do arquivo não corresponde ao formato de imagem informado.');
  }
}

function getPlanoEvidenceFolder_() {
  const properties = PropertiesService.getScriptProperties();
  const folderId = properties.getProperty(PLANO_ACAO_CONFIG.EVIDENCE_FOLDER_PROPERTY);
  if (folderId) {
    try {
      const folder = DriveApp.getFolderById(folderId);
      if (!folder.isTrashed()) return folder;
    } catch (error) {
      // A pasta será recriada se tiver sido excluída ou perdido o acesso.
    }
  }

  const spreadsheetFile = DriveApp.getFileById(getSpreadsheet_().getId());
  const parents = spreadsheetFile.getParents();
  const parent = parents.hasNext() ? parents.next() : DriveApp.getRootFolder();
  const folder = parent.createFolder('Mesa Operacional - Evidências');
  properties.setProperty(PLANO_ACAO_CONFIG.EVIDENCE_FOLDER_PROPERTY, folder.getId());
  return folder;
}

function createPlanoCreationSnapshot_(payload, user) {
  const source = payload || {};
  const city = normalizeCity_(source.cidade);
  const currentMonth = normalizeMonthKey_(source.periodo_analisado);
  const previousMonth = normalizeMonthKey_(source.periodo_comparacao);
  if (!city) throw new PlanoUserError('Informe a cidade.');
  if (!currentMonth || !previousMonth || currentMonth === previousMonth) {
    throw new PlanoUserError('Informe dois meses válidos e diferentes para o plano.');
  }

  const regional = resolvePlanoRegional_(city);
  requireAppRegionalAccess_(user, regional);
  const dashboard = getDashboardDataForUser_(user, false, currentMonth, previousMonth);
  const record = dashboard.cities.find((item) => normalizeCity_(item.city) === city);
  if (!record) {
    throw new PlanoUserError('A cidade não possui dados nos períodos selecionados ou está fora do seu acesso.');
  }
  if (normalizeAppRegional_(record.regional) !== normalizeAppRegional_(regional)) {
    throw new PlanoUserError('A regional da cidade diverge entre as bases. Atualize a planilha antes de criar o plano.');
  }

  return {
    cidade: city,
    regional,
    periodo_analisado: dashboard.period.currentKey,
    periodo_comparacao: dashboard.period.previousKey,
    base_anterior: Number(record.previous) || 0,
    base_atual: Number(record.current) || 0,
    diferenca: Number(record.difference) || 0,
    variacao_pct: Number(record.percentage) || 0,
    contexto_criacao: {
      periodo_analisado_label: dashboard.period.currentLabel,
      periodo_comparacao_label: dashboard.period.previousLabel,
      origem: 'dashboard_mesa_operacional',
    },
  };
}

function resolvePlanoRegional_(city) {
  const normalizedCity = normalizeCity_(city);
  if (!normalizedCity) throw new PlanoUserError('Informe a cidade.');

  const sheet = getRequiredSheet_(getSpreadsheet_(), CONFIG.REGION_SHEET);
  const values = sheet.getDataRange().getValues();
  const headerRow = findHeaderRow_(values, ['cidade', 'regional']);
  const headers = createHeaderMap_(values[headerRow]);
  const cityIndex = getHeaderIndex_(headers, 'cidade');
  const regionalIndex = getHeaderIndex_(headers, 'regional');
  const regionals = Array.from(new Set(
    values.slice(headerRow + 1)
      .filter((row) => normalizeCity_(row[cityIndex]) === normalizedCity)
      .map((row) => normalizeAppRegional_(row[regionalIndex]))
      .filter(Boolean)
  ));

  if (!regionals.length) {
    throw new PlanoUserError(`Cidade não cadastrada na aba ${CONFIG.REGION_SHEET}: ${normalizedCity}.`);
  }
  if (regionals.length > 1) {
    throw new PlanoUserError(`A cidade ${normalizedCity} possui regionais conflitantes na aba ${CONFIG.REGION_SHEET}.`);
  }
  return regionals[0];
}

function validatePlanoOperationId_(value) {
  const operationId = String(value || '').trim();
  if (!/^[A-Za-z0-9_-]{16,100}$/.test(operationId)) {
    throw new PlanoUserError('Identificador da operação inválido. Reabra o formulário e tente novamente.');
  }
  return operationId;
}

function findPlanoRecordByOperation_(sheetName, operationId) {
  return readPlanoRecordsByField_(sheetName, 'operacao_id', operationId)[0] || null;
}

function requireOwnedPlanoOperation_(record, user, operationId) {
  if (
    String(record.operacao_id || '') !== operationId
    || String(record.criado_por || '').trim().toLowerCase() !== user.email
  ) {
    throw new PlanoUserError('Identificador da operação já utilizado. Reabra o formulário.');
  }
  requireAccessiblePlan_(record.id, user);
}

function requireMatchingPlanoOperation_(record, planId, operationId) {
  if (
    String(record.operacao_id || '') !== operationId
    || String(record.plano_id || '') !== String(planId || '')
  ) {
    throw new PlanoUserError('Identificador da operação já utilizado. Reabra o formulário.');
  }
}

function requirePlanoUpdateBelongsToPlan_(updateId, planId) {
  const normalizedUpdateId = String(updateId || '').trim();
  if (!normalizedUpdateId) return;
  const update = findPlanoRecordById_(PLANO_ACAO_CONFIG.SHEETS.UPDATES, normalizedUpdateId);
  if (!update || String(update.plano_id || '') !== String(planId || '')) {
    throw new PlanoUserError('A atualização informada não pertence a este plano.');
  }
}

function rollbackPlanoCreate_(operationId) {
  safelyDeletePlanoRecordsByOperation_(PLANO_ACAO_CONFIG.SHEETS.HISTORY, operationId);
  safelyDeletePlanoRecordsByOperation_(PLANO_ACAO_CONFIG.SHEETS.PLANS, operationId);
}

function rollbackPlanoMutation_(existing, operationId) {
  safelyDeletePlanoRecordsByOperation_(PLANO_ACAO_CONFIG.SHEETS.HISTORY, operationId);
  safelyDeletePlanoRecordsByOperation_(PLANO_ACAO_CONFIG.SHEETS.UPDATES, operationId);
  try {
    const original = Object.assign({}, existing);
    delete original.__row;
    updatePlanoRecord_(getPlanoSheet_(PLANO_ACAO_CONFIG.SHEETS.PLANS), existing.__row, original);
    SpreadsheetApp.flush();
  } catch (rollbackError) {
    console.error(`Falha ao restaurar o plano ${existing && existing.id}: ${rollbackError.message}`);
  }
}

function safelyDeletePlanoRecordsByOperation_(sheetName, operationId) {
  if (!operationId) return;
  try {
    deletePlanoRecordsByField_(sheetName, 'operacao_id', operationId);
  } catch (rollbackError) {
    console.error(`Falha ao desfazer operação em ${sheetName}: ${rollbackError.message}`);
  }
}

function safelyDeletePlanoRow_(sheetName, rowNumber) {
  if (!Number.isInteger(Number(rowNumber)) || Number(rowNumber) < 2) return;
  try {
    deletePlanoRows_(getPlanoSheet_(sheetName), [Number(rowNumber)]);
  } catch (rollbackError) {
    console.error(`Falha ao desfazer gravação em ${sheetName}: ${rollbackError.message}`);
  }
}

function validatePlanoOption_(value, allowed, field) {
  const normalized = String(value || '').trim().toLowerCase();
  if (!allowed.includes(normalized)) throw new PlanoUserError(`Valor inválido para ${field}: ${value}`);
  return normalized;
}

function validatePlanoDate_(value) {
  const timezone = getPlanoTimeZone_();
  const text = value instanceof Date && !Number.isNaN(value.getTime())
    ? Utilities.formatDate(value, timezone, 'yyyy-MM-dd')
    : String(value || '').trim();
  if (!text) return '';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) throw new PlanoUserError('A data deve estar no formato AAAA-MM-DD.');

  // O input date envia apenas AAAA-MM-DD. Interpretá-lo com new Date(text)
  // cria meia-noite UTC e pode recuar um dia ao exibir no fuso da planilha.
  // Meio-dia no fuso da planilha preserva o dia civil mesmo com conversões.
  const date = Utilities.parseDate(`${text} 12:00:00`, timezone, 'yyyy-MM-dd HH:mm:ss');
  if (
    Number.isNaN(date.getTime())
    || Utilities.formatDate(date, timezone, 'yyyy-MM-dd') !== text
  ) {
    throw new PlanoUserError(`Data inválida: ${text}`);
  }
  return date;
}

function clampPlanoPercentage_(value) {
  const number = Number(value);
  if (!Number.isFinite(number) || number < 0 || number > 100) {
    throw new PlanoUserError('O percentual de conclusão deve estar entre 0 e 100.');
  }
  return Math.round(number);
}

function planoNumber_(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : 0;
}

function planoLimitedText_(value, maximumLength, fieldLabel) {
  const text = String(value || '').trim();
  if (text.length > maximumLength) {
    throw new PlanoUserError(`${fieldLabel} deve ter no máximo ${maximumLength} caracteres.`);
  }
  return text;
}

function normalizePlanoLinks_(value) {
  const parsed = Array.isArray(value) ? value : parsePlanoJson_(value, null);
  const links = (Array.isArray(parsed) ? parsed : String(value || '').split(/[\n,]+/))
    .map((link) => String(link || '').trim())
    .filter(Boolean);
  if (links.length > PLANO_ACAO_CONFIG.MAX_LINKS) {
    throw new PlanoUserError(`Informe no máximo ${PLANO_ACAO_CONFIG.MAX_LINKS} links.`);
  }

  links.forEach((link) => {
    if (link.length > PLANO_ACAO_CONFIG.MAX_LINK_LENGTH) {
      throw new PlanoUserError(`Cada link deve ter no máximo ${PLANO_ACAO_CONFIG.MAX_LINK_LENGTH} caracteres.`);
    }
    if (!/^https?:\/\//i.test(link)) {
      throw new PlanoUserError('Os links devem começar com http:// ou https://.');
    }
  });
  return Array.from(new Set(links));
}

function comparePlanoDatesDesc_(left, right) {
  const leftTime = new Date(left.atualizado_em || left.criado_em || 0).getTime() || 0;
  const rightTime = new Date(right.atualizado_em || right.criado_em || 0).getTime() || 0;
  return rightTime - leftTime;
}

function planoStatusLabel_(status) {
  return {
    nao_iniciado: 'Não iniciado',
    em_andamento: 'Em andamento',
    pendente: 'Pendente',
    concluido: 'Concluído',
    cancelado: 'Cancelado',
  }[status] || status;
}

function planoPriorityLabel_(priority) {
  return {
    baixa: 'Baixa',
    media: 'Média',
    alta: 'Alta',
    critica: 'Crítica',
  }[priority] || priority;
}

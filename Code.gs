/**
 * ============================================================================
 *  BACKEND DO PORTFÓLIO — Sheilla Soares Arquitetura & Design
 *  Google Apps Script (Web App) + Planilha Google + Google Drive
 * ============================================================================
 *
 *  O que este script faz:
 *   - Guarda os dados dos projetos em uma planilha (criada automaticamente).
 *   - Guarda as fotos em uma pasta do Google Drive (criada automaticamente).
 *   - Valida a senha NO SERVIDOR e só libera cadastrar/editar/excluir com sessão válida.
 *   - Entrega a lista pública de projetos para o site.
 *
 *  A SENHA NÃO FICA NESTE CÓDIGO. Ela fica em:
 *   Configurações do projeto (engrenagem) > Propriedades do script > ADMIN_PASSWORD
 *
 *  Passo a passo completo: veja o arquivo LEIA-ME.md.
 * ============================================================================
 */

var CFG = {
  SHEET_NAME: 'Projetos',
  SS_TITLE: 'Portfólio Sheilla Soares - Dados',
  FOLDER_NAME: 'Portfólio Sheilla Soares - Fotos',
  HEADERS: ['id', 'titulo', 'descricao', 'categoria', 'localizacao', 'capa', 'ordem', 'fotos', 'criadoEm', 'atualizadoEm'],
  CATEGORIES: ['Interior', 'Exterior', 'Projeto Arquitetônico'],
  SESSION_TTL: 21600,            // 6 horas (máximo do CacheService); renova a cada uso
  MAX_FAILS: 5,                  // tentativas de senha erradas antes de bloquear
  LOCK_SECONDS: 300,             // tempo de bloqueio (5 min)
  MAX_PHOTO_BYTES: 6 * 1024 * 1024,
  MAX_PHOTOS_PER_PROJECT: 40,
  MAX_TITLE: 120,
  MAX_LOCATION: 120,
  MAX_DESC: 3000,
  PUBLIC_CACHE_KEY: 'pub_list_v1',
  PUBLIC_CACHE_TTL: 600
};

/* ============================================================================
 *  CONFIGURAÇÃO INICIAL (rodar UMA vez, no editor)
 * ========================================================================== */

function configurar() {
  var props = PropertiesService.getScriptProperties();

  // Planilha
  var ss = null;
  var ssId = props.getProperty('SS_ID');
  if (ssId) {
    try { ss = SpreadsheetApp.openById(ssId); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(CFG.SS_TITLE);
    props.setProperty('SS_ID', ss.getId());
  }
  var sh = ss.getSheetByName(CFG.SHEET_NAME);
  if (!sh) {
    var first = ss.getSheets()[0];
    if (first && first.getLastRow() === 0 && ss.getSheets().length === 1) {
      sh = first;
      sh.setName(CFG.SHEET_NAME);
    } else {
      sh = ss.insertSheet(CFG.SHEET_NAME);
    }
  }
  sh.getRange(1, 1, 1, CFG.HEADERS.length).setValues([CFG.HEADERS]);
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1000, CFG.HEADERS.length).setNumberFormat('@'); // tudo como texto (evita fórmulas)

  // Pasta de fotos
  var folder = null;
  var folderId = props.getProperty('FOLDER_ID');
  if (folderId) {
    try { folder = DriveApp.getFolderById(folderId); } catch (e) { folder = null; }
  }
  if (!folder) {
    folder = DriveApp.createFolder(CFG.FOLDER_NAME);
    props.setProperty('FOLDER_ID', folder.getId());
  }
  try {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {
    Logger.log('Aviso: não foi possível tornar a pasta pública (' + e + '). ' +
      'Se a conta tiver restrição de compartilhamento, as fotos não aparecerão no site.');
  }

  var temSenha = !!props.getProperty('ADMIN_PASSWORD');
  Logger.log('Planilha: ' + ss.getUrl());
  Logger.log('Pasta de fotos: ' + folder.getUrl());
  Logger.log(temSenha
    ? 'Senha de administrador: configurada.'
    : 'ATENÇÃO: falta criar a propriedade ADMIN_PASSWORD em Configurações do projeto > Propriedades do script.');
  Logger.log('Configuração concluída. Agora faça a implantação como App da Web (veja o LEIA-ME).');
}

/**
 * Opcional, rode manualmente de vez em quando: apaga (envia à lixeira do Drive)
 * fotos que ficaram soltas por uploads interrompidos e não pertencem a nenhum projeto.
 */
function limparFotosOrfas() {
  var usadas = {};
  readAll_().forEach(function (p) { p.fotos.forEach(function (id) { usadas[id] = true; }); });
  var umDia = 24 * 60 * 60 * 1000;
  var it = getFolder_().getFiles();
  var removidas = 0;
  while (it.hasNext()) {
    var f = it.next();
    if (!usadas[f.getId()] && (Date.now() - f.getDateCreated().getTime()) > umDia) {
      f.setTrashed(true);
      removidas++;
    }
  }
  Logger.log('Fotos órfãs enviadas para a lixeira: ' + removidas);
}

/* ============================================================================
 *  ENTRADAS HTTP
 * ========================================================================== */

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'ping';
    if (action === 'list') return json_(publicPayload_());
    return json_({ ok: true, service: 'portfolio' });
  } catch (err) {
    return json_(errorPayload_(err));
  }
}

function doPost(e) {
  try {
    var body = parseBody_(e);
    var action = String(body.action || '');

    if (action === 'list') return json_(publicPayload_());
    if (action === 'login') return json_(login_(body.password));

    // Tudo abaixo exige sessão válida (validada aqui, no servidor)
    if (!checkSession_(body.token)) {
      throw fail_('sessao_invalida', 'Sua sessão expirou. Entre novamente com a senha.');
    }

    switch (action) {
      case 'logout':  return json_(logout_(body.token));
      case 'upload':  return json_(uploadPhoto_(body));
      case 'save':    return json_(saveProject_(body));
      case 'delete':  return json_(deleteProject_(body));
      case 'reorder': return json_(reorderProjects_(body));
      case 'discard': return json_(discardPhotos_(body));
      default: throw fail_('acao_invalida', 'Ação desconhecida.');
    }
  } catch (err) {
    return json_(errorPayload_(err));
  }
}

/* ============================================================================
 *  AUTENTICAÇÃO
 * ========================================================================== */

function login_(password) {
  var cache = CacheService.getScriptCache();
  var fails = Number(cache.get('login_fails') || 0);
  if (fails >= CFG.MAX_FAILS) {
    throw fail_('bloqueado', 'Muitas tentativas incorretas. Aguarde alguns minutos e tente de novo.');
  }
  var expected = PropertiesService.getScriptProperties().getProperty('ADMIN_PASSWORD');
  if (!expected) {
    throw fail_('sem_senha', 'A senha de administrador ainda não foi configurada no servidor.');
  }
  var typed = typeof password === 'string' ? password : '';
  if (!typed || !safeEqual_(hash_(typed), hash_(expected))) {
    cache.put('login_fails', String(fails + 1), CFG.LOCK_SECONDS);
    Utilities.sleep(800); // desacelera tentativas automáticas
    throw fail_('senha_incorreta', 'Senha incorreta. Confira e tente novamente.');
  }
  cache.remove('login_fails');
  var token = (Utilities.getUuid() + Utilities.getUuid()).replace(/-/g, '');
  cache.put('s_' + hash_(token), '1', CFG.SESSION_TTL);
  return { ok: true, token: token };
}

function checkSession_(token) {
  if (typeof token !== 'string' || token.length < 32) return false;
  var cache = CacheService.getScriptCache();
  var key = 's_' + hash_(token);
  if (cache.get(key)) {
    cache.put(key, '1', CFG.SESSION_TTL); // renova enquanto estiver em uso
    return true;
  }
  return false;
}

function logout_(token) {
  CacheService.getScriptCache().remove('s_' + hash_(token));
  return { ok: true };
}

function hash_(s) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s), Utilities.Charset.UTF_8);
  return bytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');
}

function safeEqual_(a, b) {
  if (a.length !== b.length) return false;
  var diff = 0;
  for (var i = 0; i < a.length; i++) diff |= a.charCodeAt(i) ^ b.charCodeAt(i);
  return diff === 0;
}

/* ============================================================================
 *  LEITURA PÚBLICA
 * ========================================================================== */

function publicPayload_() {
  var cache = CacheService.getScriptCache();
  var hit = cache.get(CFG.PUBLIC_CACHE_KEY);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) { /* segue e recalcula */ }
  }
  var payload = {
    ok: true,
    projects: readAll_().map(toPublic_),
    hadProjects: PropertiesService.getScriptProperties().getProperty('HAD_PROJECTS') === '1'
  };
  try { cache.put(CFG.PUBLIC_CACHE_KEY, JSON.stringify(payload), CFG.PUBLIC_CACHE_TTL); } catch (e) { /* grande demais para o cache: ok */ }
  return payload;
}

function invalidatePublicCache_() {
  CacheService.getScriptCache().remove(CFG.PUBLIC_CACHE_KEY);
}

function toPublic_(p) {
  return {
    id: p.id,
    titulo: p.titulo,
    descricao: p.descricao,
    categoria: p.categoria,
    localizacao: p.localizacao,
    capa: p.capa,
    ordem: p.ordem,
    fotos: p.fotos
  };
}

/* ============================================================================
 *  UPLOAD DE FOTO
 * ========================================================================== */

function uploadPhoto_(body) {
  var data = typeof body.data === 'string' ? body.data.replace(/^data:[^;]+;base64,/, '') : '';
  if (!data) throw fail_('arquivo_invalido', 'Nenhuma imagem recebida.');
  if (data.length * 0.75 > CFG.MAX_PHOTO_BYTES) {
    throw fail_('arquivo_grande', 'A imagem é grande demais (máximo de 6 MB após a otimização).');
  }
  var bytes;
  try { bytes = Utilities.base64Decode(data); } catch (e) { throw fail_('arquivo_invalido', 'Imagem corrompida.'); }
  if (bytes.length > CFG.MAX_PHOTO_BYTES) throw fail_('arquivo_grande', 'A imagem é grande demais.');
  var mime = detectMime_(bytes);
  if (!mime) throw fail_('formato_invalido', 'Formato não aceito. Envie JPG, PNG ou WEBP.');

  var ext = mime === 'image/jpeg' ? 'jpg' : (mime === 'image/png' ? 'png' : 'webp');
  var name = 'foto-' + Date.now() + '-' + Utilities.getUuid().slice(0, 8) + '.' + ext;
  var file = getFolder_().createFile(Utilities.newBlob(bytes, mime, name));
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) { /* herda da pasta */ }
  return { ok: true, id: file.getId() };
}

/** Confere o tipo real do arquivo pelos primeiros bytes (não confia no que o navegador diz). */
function detectMime_(b) {
  if (b.length > 12) {
    if (b[0] === -1 && b[1] === -40 && b[2] === -1) return 'image/jpeg';
    if (b[0] === -119 && b[1] === 0x50 && b[2] === 0x4E && b[3] === 0x47) return 'image/png';
    if (b[0] === 0x52 && b[1] === 0x49 && b[2] === 0x46 && b[3] === 0x46 &&
        b[8] === 0x57 && b[9] === 0x45 && b[10] === 0x42 && b[11] === 0x50) return 'image/webp';
  }
  return null;
}

/* ============================================================================
 *  CRUD DE PROJETOS
 * ========================================================================== */

function saveProject_(body) {
  var titulo = cleanText_(body.titulo, CFG.MAX_TITLE);
  var categoria = String(body.categoria || '');
  var localizacao = cleanText_(body.localizacao, CFG.MAX_LOCATION);
  var descricao = cleanText_(body.descricao, CFG.MAX_DESC, true);
  var fotos = Array.isArray(body.fotos) ? body.fotos.map(String) : [];

  if (!titulo) throw fail_('titulo_obrigatorio', 'Informe o título do projeto.');
  if (CFG.CATEGORIES.indexOf(categoria) < 0) throw fail_('categoria_invalida', 'Escolha a categoria: Interior, Exterior ou Projeto Arquitetônico.');
  if (fotos.length > CFG.MAX_PHOTOS_PER_PROJECT) {
    throw fail_('muitas_fotos', 'Máximo de ' + CFG.MAX_PHOTOS_PER_PROJECT + ' fotos por projeto.');
  }
  fotos = fotos.filter(function (id, i) { return /^[A-Za-z0-9_-]{10,}$/.test(id) && fotos.indexOf(id) === i; });

  return withLock_(function () {
    var all = readAll_();
    var existing = null;
    if (body.id) {
      existing = all.filter(function (p) { return p.id === String(body.id); })[0] || null;
      if (!existing) throw fail_('nao_encontrado', 'Projeto não encontrado. Atualize a página e tente de novo.');
    }
    var oldFotos = existing ? existing.fotos : [];

    // Só aceita fotos novas que realmente estejam na nossa pasta
    fotos.forEach(function (id) {
      if (oldFotos.indexOf(id) < 0 && !isOurFile_(id)) {
        throw fail_('foto_invalida', 'Uma das fotos não foi enviada corretamente. Tente enviar de novo.');
      }
    });

    var capa = String(body.capa || '');
    if (fotos.indexOf(capa) < 0) capa = fotos[0] || '';

    var now = new Date().toISOString();
    var proj = {
      id: existing ? existing.id : 'p' + Utilities.getUuid().replace(/-/g, '').slice(0, 12),
      titulo: titulo,
      descricao: descricao,
      categoria: categoria,
      localizacao: localizacao,
      capa: capa,
      ordem: 0,
      fotos: fotos,
      criadoEm: existing ? existing.criadoEm : now,
      atualizadoEm: now
    };

    // Posição na lista (1 = primeiro). Padrão: mantém a atual; novo vai para o fim.
    var rest = all.filter(function (p) { return !existing || p.id !== existing.id; });
    var idx = existing ? all.indexOf(existing) : rest.length;
    var pos = parseInt(body.posicao, 10);
    if (pos >= 1) idx = Math.min(pos - 1, rest.length);
    rest.splice(idx, 0, proj);
    writeAll_(rest);

    // Fotos removidas do projeto vão para a lixeira do Drive (recuperáveis por 30 dias)
    oldFotos.forEach(function (id) { if (fotos.indexOf(id) < 0) trashPhoto_(id); });

    PropertiesService.getScriptProperties().setProperty('HAD_PROJECTS', '1');
    invalidatePublicCache_();
    return { ok: true, id: proj.id, projects: readAll_().map(toPublic_) };
  });
}

function deleteProject_(body) {
  return withLock_(function () {
    var all = readAll_();
    var target = all.filter(function (p) { return p.id === String(body.id); })[0];
    if (!target) throw fail_('nao_encontrado', 'Projeto não encontrado. Atualize a página.');
    writeAll_(all.filter(function (p) { return p !== target; }));
    target.fotos.forEach(trashPhoto_);
    invalidatePublicCache_();
    return { ok: true, projects: readAll_().map(toPublic_) };
  });
}

function reorderProjects_(body) {
  var ids = Array.isArray(body.ids) ? body.ids.map(String) : [];
  return withLock_(function () {
    var all = readAll_();
    var byId = {};
    all.forEach(function (p) { byId[p.id] = p; });
    var ordered = [];
    ids.forEach(function (id) { if (byId[id]) { ordered.push(byId[id]); delete byId[id]; } });
    all.forEach(function (p) { if (byId[p.id]) ordered.push(p); }); // o que faltou vai para o fim
    writeAll_(ordered);
    invalidatePublicCache_();
    return { ok: true, projects: readAll_().map(toPublic_) };
  });
}

/** Remove fotos enviadas mas não salvas (usuário cancelou o formulário). Só mexe em fotos órfãs da nossa pasta. */
function discardPhotos_(body) {
  var ids = Array.isArray(body.ids) ? body.ids.map(String).slice(0, CFG.MAX_PHOTOS_PER_PROJECT) : [];
  var usadas = {};
  readAll_().forEach(function (p) { p.fotos.forEach(function (id) { usadas[id] = true; }); });
  var n = 0;
  ids.forEach(function (id) {
    if (!usadas[id] && trashPhoto_(id)) n++;
  });
  return { ok: true, removed: n };
}

/* ============================================================================
 *  PLANILHA E DRIVE
 * ========================================================================== */

function getSheet_() {
  var id = PropertiesService.getScriptProperties().getProperty('SS_ID');
  if (!id) throw fail_('nao_configurado', 'O servidor ainda não foi configurado. Execute a função "configurar" no editor do Apps Script.');
  var sh = SpreadsheetApp.openById(id).getSheetByName(CFG.SHEET_NAME);
  if (!sh) throw fail_('nao_configurado', 'Aba "' + CFG.SHEET_NAME + '" não encontrada. Execute "configurar" novamente.');
  return sh;
}

function getFolder_() {
  var id = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
  if (!id) throw fail_('nao_configurado', 'A pasta de fotos ainda não foi criada. Execute a função "configurar".');
  return DriveApp.getFolderById(id);
}

function readAll_() {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last < 2) return [];
  var rows = sh.getRange(2, 1, last - 1, CFG.HEADERS.length).getValues();
  var list = [];
  rows.forEach(function (r) {
    if (!r[0]) return;
    var fotos = [];
    try { fotos = JSON.parse(r[7] || '[]'); } catch (e) { fotos = []; }
    list.push({
      id: String(r[0]),
      titulo: String(r[1] || ''),
      descricao: String(r[2] || ''),
      categoria: String(r[3] || ''),
      localizacao: String(r[4] || ''),
      capa: String(r[5] || ''),
      ordem: Number(r[6]) || 0,
      fotos: Array.isArray(fotos) ? fotos.map(String) : [],
      criadoEm: String(r[8] || ''),
      atualizadoEm: String(r[9] || '')
    });
  });
  list.sort(function (a, b) { return a.ordem - b.ordem; });
  return list;
}

/** Regrava a tabela inteira, já renumerando a ordem (1..n). */
function writeAll_(list) {
  var sh = getSheet_();
  var last = sh.getLastRow();
  if (last >= 2) sh.getRange(2, 1, last - 1, CFG.HEADERS.length).clearContent();
  if (!list.length) return;
  var values = list.map(function (p, i) {
    p.ordem = i + 1;
    return [p.id, p.titulo, p.descricao, p.categoria, p.localizacao, p.capa, p.ordem,
            JSON.stringify(p.fotos), p.criadoEm, p.atualizadoEm];
  });
  var range = sh.getRange(2, 1, values.length, CFG.HEADERS.length);
  range.setNumberFormat('@');
  range.setValues(values);
}

function isOurFile_(id) {
  try {
    var folderId = PropertiesService.getScriptProperties().getProperty('FOLDER_ID');
    var parents = DriveApp.getFileById(id).getParents();
    while (parents.hasNext()) {
      if (parents.next().getId() === folderId) return true;
    }
  } catch (e) { /* arquivo inexistente ou sem acesso */ }
  return false;
}

function trashPhoto_(id) {
  if (!isOurFile_(id)) return false;
  try { DriveApp.getFileById(id).setTrashed(true); return true; } catch (e) { return false; }
}

/* ============================================================================
 *  UTILITÁRIOS
 * ========================================================================== */

function withLock_(fn) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
  } catch (e) {
    throw fail_('ocupado', 'O servidor está ocupado. Tente novamente em instantes.');
  }
  try { return fn(); } finally { lock.releaseLock(); }
}

function cleanText_(v, max, keepNewlines) {
  var s = String(v == null ? '' : v).replace(/\r\n?/g, '\n');
  s = keepNewlines ? s.replace(/[^\S\n]+/g, ' ').replace(/\n{3,}/g, '\n\n') : s.replace(/\s+/g, ' ');
  return s.trim().slice(0, max);
}

function parseBody_(e) {
  try {
    return JSON.parse((e && e.postData && e.postData.contents) || '{}') || {};
  } catch (err) {
    throw fail_('requisicao_invalida', 'Requisição inválida.');
  }
}

function fail_(code, message) {
  var err = new Error(message);
  err.code = code;
  err.isPublic = true;
  return err;
}

function errorPayload_(err) {
  if (err && err.isPublic) return { ok: false, error: err.code, message: err.message };
  Logger.log('Erro interno: ' + (err && err.stack ? err.stack : err));
  return { ok: false, error: 'erro_interno', message: 'Ocorreu um erro no servidor. Tente novamente.' };
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

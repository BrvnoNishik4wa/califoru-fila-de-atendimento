const SHEET_ID        = 'ID';   // ← ID da planilha de dados
const DRIVE_FOLDER_ID = 'ID';        // ← ID da pasta do Drive para os relatórios exportados

// ── WHITELIST ─────────────────────────────────────────────
// ÚNICA fonte de verdade. A planilha não interfere aqui.
// Adicione ou remova e-mails diretamente neste objeto.
const WHITELIST_LDV = {
  'loja1@empresa-exemplo.com':  { role: 'store', storeKey: 'loja1', storeName: 'Loja 1' },
  'loja2@empresa-exemplo.com':  { role: 'store', storeKey: 'loja2', storeName: 'Loja 2' },
  'gestora@empresa-exemplo.com': { role: 'admin', storeKey: '',     storeName: 'Gestora' },
};

// ── ABAS ──────────────────────────────────────────────────
const ABA = {
  FILA:  'fila_estado',
  ATEND: 'atendimentos',
  VEND:  'vendedores',
};

// Índices das colunas da aba atendimentos (base 1)
const COL = {
  ID: 1, TS: 2, DATA: 3, HORA: 4,
  LOJA_KEY: 5, LOJA_NOME: 6, VEND: 7,
  RESULTADO: 8, MOTIVO: 9, IS_TROCA: 10,
  PRODUTO_ESP: 11,
};

// Motivo de perda que abre o campo de texto livre "que produto o cliente
// procurava". O texto vai para COL.PRODUTO_ESP — NUNCA para COL.MOTIVO,
// que precisa continuar sendo um conjunto fechado de strings para as
// agregações de motivos seguirem funcionando.
const MOTIVO_PRODUTO_ESP = 'Produto específico';

// ── AUSÊNCIAS MOMENTÂNEAS (REGRAS) ────────────────────────
// Só o banheiro guarda a posição na fila. Pessoal e intervalo mandam o
// vendedor para o fim ao voltar. Guardado no JSON de fila_estado, no campo
// tipoPausa do item — só existe enquanto status === 'pausa'.
const PAUSA_TIPOS = ['banheiro', 'pessoal', 'intervalo'];
const PAUSA_MANTEM_POSICAO = 'banheiro';

// ── ENTRY POINTS ──────────────────────────────────────────
function doGet() {
  return HtmlService
    .createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Lista da Vez — CaliforU')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ── DEBUG — execute pelo editor para ver o e-mail detectado ─
function debugEmail() {
  const email = Session.getActiveUser().getEmail();
  Logger.log('Email detectado: ' + email);
  Logger.log('Na whitelist: ' + (WHITELIST_LDV[email.trim().toLowerCase()] ? 'SIM ✅' : 'NÃO ❌'));
}

// Execute esta função UMA VEZ pelo editor para autorizar o DriveApp
function autorizarDrive() {
  const test = DriveApp.getRootFolder().getName();
  Logger.log('Drive autorizado ✅ — pasta raiz: ' + test);
}

// ── SETUP INICIAL ─────────────────────────────────────────
// Execute UMA VEZ para criar as abas necessárias
function setupPlanilha() {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  // fila_estado
  let aba = ss.getSheetByName(ABA.FILA) || ss.insertSheet(ABA.FILA);
  if (!aba.getRange('A1').getValue()) {
    aba.getRange('A1').setValue('loja2');
    aba.getRange('B1').setValue('loja1');
    aba.getRange('A2').setValue('[]');
    aba.getRange('B2').setValue('[]');
  }

  // atendimentos
  aba = ss.getSheetByName(ABA.ATEND) || ss.insertSheet(ABA.ATEND);
  if (aba.getLastRow() === 0) {
    aba.appendRow(['ID','Timestamp ISO','Data','Hora','Loja Key','Loja Nome','Vendedor','Resultado','Motivo Perda','Troca/Ajuste','Produto Específico']);
    aba.getRange(1,1,1,11).setFontWeight('bold').setBackground('#1a1a1a').setFontColor('#ffffff');
    aba.setFrozenRows(1);
  }

  // vendedores
  aba = ss.getSheetByName(ABA.VEND) || ss.insertSheet(ABA.VEND);
  if (aba.getLastRow() === 0) {
    aba.appendRow(['loja_key','loja_nome','vendedor','ativo']);
    aba.getRange(1,1,1,4).setFontWeight('bold').setBackground('#1a1a1a').setFontColor('#ffffff');
    aba.setFrozenRows(1);
    aba.appendRow(['loja2','Loja 2',  'Camila',      'TRUE']);
    aba.appendRow(['loja2','Loja 2',  'Bianca',     'TRUE']);
    aba.appendRow(['loja2','Loja 2',  'Renata',        'TRUE']);
    aba.appendRow(['loja1','Loja 1', 'Rafael',      'TRUE']);
    aba.appendRow(['loja1','Loja 1', 'Vinicius',       'TRUE']);
    aba.appendRow(['loja1','Loja 1', 'Patrícia Alves','TRUE']);
  }

  // Remove aba config se existir — não é mais usada
  const abaConfig = ss.getSheetByName('config');
  if (abaConfig) {
    ss.deleteSheet(abaConfig);
    Logger.log('Aba "config" removida com sucesso.');
  }

  SpreadsheetApp.getUi().alert('✅ Planilha configurada! Aba "config" removida.');
}

// ── MIGRAÇÃO PONTUAL ──────────────────────────────────────
// Execute UMA VEZ pelo editor na planilha que já está em produção.
// Só escreve o cabeçalho da coluna 11 — nenhuma linha existente é tocada
// (registros antigos ficam com a célula vazia, que é o valor correto para
// eles). Rodar de novo por engano é inofensivo: a função detecta que a
// coluna já existe e sai sem alterar nada.
function migrarColunaProdutoEspecifico() {
  const aba = getAba(ABA.ATEND);
  if (!aba) throw new Error('Aba "' + ABA.ATEND + '" não encontrada.');

  // Garante que a planilha tem colunas suficientes antes de escrever
  const faltam = COL.PRODUTO_ESP - aba.getMaxColumns();
  if (faltam > 0) aba.insertColumnsAfter(aba.getMaxColumns(), faltam);

  const cel   = aba.getRange(1, COL.PRODUTO_ESP);
  const atual = String(cel.getValue()).trim();

  if (atual === 'Produto Específico') {
    Logger.log('Coluna já migrada — nada a fazer.');
    return;
  }
  if (atual) {
    throw new Error('A coluna ' + COL.PRODUTO_ESP + ' já contém "' + atual +
                    '". Migração abortada para não sobrescrever dados.');
  }

  cel.setValue('Produto Específico')
     .setFontWeight('bold').setBackground('#1a1a1a').setFontColor('#ffffff');
  SpreadsheetApp.flush();
  Logger.log('✅ Coluna "Produto Específico" criada em ' + ABA.ATEND +
             '. Linhas existentes não foram alteradas.');
}

// ── RESET DIÁRIO DA FILA ──────────────────────────────────
// A fila nunca reiniciava sozinha: se ninguém saísse manualmente no fim do
// expediente, ela reaparecia idêntica na manhã seguinte, com os horários de
// ontem. E a tela mostra só a HORA no "NA FILA DESDE", nunca a data — então
// o time não tinha como perceber que estava olhando a fila do dia anterior,
// nem confiar na ordem dela. Zerar à meia-noite corrige na origem.
//
// Roda pelo gatilho de tempo do sistema, NÃO por google.script.run: não
// existe usuário na sessão para validar, por isso aqui não entra _sessao()
// nem _requireAdmin(). Instalação em instalarTriggerResetDiario(), abaixo.
function resetFilasDiario() {
  // Mesmo lock das outras escritas de fila: à meia-noite ninguém deve estar
  // usando, mas se estiver, evita gravar em cima de uma mutação em curso.
  // Se o lock não vier, a execução falha e aparece em Execuções — visível,
  // em vez de uma escrita pela metade.
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const cols  = _getFilaCols();          // uma coluna por loja da aba
    const lojas = Object.keys(cols);

    // Escreve pelo _writeFila() de propósito: assim o formato gravado é
    // exatamente o mesmo do resto do sistema (array JSON puro, '[]'), sem
    // risco de divergir se aquele formato mudar um dia.
    lojas.forEach(key => _writeFila(key, []));
    SpreadsheetApp.flush();

    const quando = Utilities.formatDate(new Date(), Session.getScriptTimeZone(),
                                        'dd/MM/yyyy HH:mm:ss');
    Logger.log('resetFilasDiario OK — ' + quando + ' — ' + lojas.length +
               ' loja(s) zerada(s): ' + (lojas.join(', ') || '(nenhuma)'));
  } finally { lock.releaseLock(); }
}

// Execute UMA VEZ pelo editor (Executar > instalarTriggerResetDiario), do
// mesmo jeito que setupPlanilha() e migrarColunaProdutoEspecifico().
// Rodar de novo por engano é inofensivo: apaga o gatilho antigo antes de
// criar o novo, porque dois gatilhos iguais significariam dois resets no
// mesmo dia.
function instalarTriggerResetDiario() {
  let removidos = 0;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'resetFilasDiario') {
      ScriptApp.deleteTrigger(t);
      removidos++;
    }
  });

  ScriptApp.newTrigger('resetFilasDiario')
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .create();

  const msg = '✅ Gatilho instalado! A fila das lojas zera todos os dias à meia-noite.' +
              (removidos ? '\n\n' + removidos + ' gatilho(s) antigo(s) removido(s) para não duplicar.' : '');

  // Loga ANTES do alert: em script standalone o getUi() pode não estar
  // disponível, e nesse caso o gatilho já está criado — a confirmação
  // precisa sobreviver no log mesmo se a caixa de diálogo não abrir.
  Logger.log(msg);
  try {
    SpreadsheetApp.getUi().alert(msg);
  } catch (err) {
    Logger.log('(Sem interface para exibir o alert — gatilho instalado de qualquer forma.)');
  }
}

// ── HELPERS ───────────────────────────────────────────────
function getSS()   { return SpreadsheetApp.openById(SHEET_ID); }
function getAba(n) { return getSS().getSheetByName(n); }

// Lê vendedores APENAS da aba vendedores (não da config)
function _loadVendedores() {
  const aba = getAba(ABA.VEND);
  if (!aba) return {};
  const rows = aba.getDataRange().getValues().slice(1);
  const map  = {};
  rows.forEach(r => {
    const [lojaKey,, vend, ativo] = r;
    if (String(ativo).toUpperCase() !== 'TRUE') return;
    const k = String(lojaKey).trim();
    if (!map[k]) map[k] = [];
    map[k].push(String(vend).trim());
  });
  return map;
}

// Deriva lojas diretamente da WHITELIST_LDV (sem planilha)
function _getLojas() {
  const lojas = [];
  const seen  = new Set();
  Object.values(WHITELIST_LDV).forEach(u => {
    if (u.role === 'store' && !seen.has(u.storeKey)) {
      seen.add(u.storeKey);
      lojas.push({ key: u.storeKey, name: u.storeName });
    }
  });
  return lojas;
}

// ── AUTH ──────────────────────────────────────────────────
// Toda função top-level de um .gs fica exposta a google.script.run. Checar
// o e-mail só no verificarAcesso() não protegia nada: qualquer conta Google
// logada podia chamar as outras direto pelo console e pular a whitelist e o
// papel (admin × store). Por isso cada função pública abaixo começa por um
// destes guards, e nenhuma confia em loja/vendedor vindos do cliente.

// Identidade de quem está chamando. Lança se não estiver na whitelist —
// erro em vez de retorno vazio para nenhuma função seguir executando por
// engano se esquecerem de checar o resultado.
function _sessao() {
  const email  = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
  const perfil = WHITELIST_LDV[email];
  if (!perfil) throw new Error('Sem permissão.');
  return {
    email,
    role:      perfil.role,
    storeKey:  perfil.storeKey,
    storeName: perfil.storeName,
  };
}

function _requireAdmin() {
  const sessao = _sessao();
  if (sessao.role !== 'admin') throw new Error('Ação restrita à gestora.');
  return sessao;
}

// Resolve a loja no servidor e devolve os valores VALIDADOS. O storeKey que
// vem do cliente é tratado como pedido, não como verdade: a gestora pode
// operar qualquer loja conhecida, a loja só pode operar a própria. Quem
// chama deve usar o retorno daqui, nunca o argumento cru.
function _requireLoja(storeKeyRecebido) {
  const sessao = _sessao();
  const key    = String(storeKeyRecebido || '').trim();

  if (sessao.role === 'admin') {
    const loja = _getLojas().find(l => l.key === key);
    if (!loja) throw new Error('Loja inválida.');
    return { ...sessao, storeKey: loja.key, storeName: loja.name };
  }

  if (key !== sessao.storeKey) throw new Error('Ação restrita à sua loja.');
  return sessao;
}

// Impede agir em nome de vendedor que não existe ou é de outra loja — sem
// isto dava para injetar qualquer nome na fila e no histórico de vendas.
function _requireVendedorDaLoja(storeKey, vend) {
  const nome  = String(vend || '').trim();
  const lista = _loadVendedores()[storeKey] || [];
  if (!nome || lista.indexOf(nome) === -1) {
    throw new Error('Vendedor não encontrado nesta loja.');
  }
  return nome;
}

// Mesma checagem de sempre, agora em cima do _sessao(). O try/catch mantém
// o contrato { ok:false, msg } que a tela de acesso negado espera — esta é
// a única função que responde "não autorizado" em vez de lançar.
function verificarAcesso() {
  let sessao;
  try {
    sessao = _sessao();
  } catch (err) {
    const email = String(Session.getActiveUser().getEmail() || '').trim().toLowerCase();
    return { ok: false, msg: 'E-mail detectado: ' + (email || '(vazio)') + ' — sem permissão.' };
  }

  const vendedores = _loadVendedores();
  const config = {
    vendedores: sessao.role === 'store'
      ? (vendedores[sessao.storeKey] || [])
      : vendedores,
    lojas: _getLojas(),
  };

  return {
    ok: true,
    user: {
      role:      sessao.role,
      storeKey:  sessao.storeKey,
      storeName: sessao.storeName,
      email:     sessao.email,
    },
    config,
  };
}

// ── FILA ──────────────────────────────────────────────────
function _getFilaCols() {
  const aba  = getAba(ABA.FILA);
  const row1 = aba.getRange(1, 1, 1, Math.max(aba.getLastColumn(), 1)).getValues()[0];
  const map  = {};
  row1.forEach((v, i) => { if (v) map[String(v).trim()] = i + 1; });
  return map;
}

function _readFila(storeKey) {
  const col = _getFilaCols()[storeKey];
  if (!col) return [];
  const val = getAba(ABA.FILA).getRange(2, col).getValue();
  try {
    const fila = JSON.parse(val) || [];
    // Garante que todo item tem status definido
    return fila.map(f => ({ ...f, status: f.status || 'aguardando' }));
  } catch { return []; }
}

function _writeFila(storeKey, fila) {
  const col = _getFilaCols()[storeKey];
  if (!col) return;
  getAba(ABA.FILA).getRange(2, col).setValue(JSON.stringify(fila));
}

function _readTodasFilas() {
  const cols  = _getFilaCols();
  const aba   = getAba(ABA.FILA);
  const filas = {};
  Object.entries(cols).forEach(([key, col]) => {
    const val = aba.getRange(2, col).getValue();
    try { filas[key] = JSON.parse(val) || []; } catch { filas[key] = []; }
  });
  return filas;
}

// ── ATENDIMENTOS ──────────────────────────────────────────
function _appendAtendimento(storeKey, storeName, vend, resultado, motivo, isTroca, produtoEspecifico) {
  const now  = new Date();
  const data = Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  const hora = Utilities.formatDate(now, Session.getScriptTimeZone(), 'HH:mm');
  getAba(ABA.ATEND).appendRow([
    now.getTime(), now.toISOString(), data, hora,
    storeKey, storeName, vend,
    resultado, motivo || '', isTroca ? 'TRUE' : 'FALSE',
    String(produtoEspecifico || '').trim(),
  ]);
  SpreadsheetApp.flush();
}

// Lê as linhas de dados da aba atendimentos com todas as colunas de COL.
// Se a planilha ainda não passou por migrarColunaProdutoEspecifico(), lê o
// que existe e completa com vazio — assim o sistema continua funcionando
// antes da migração.
function _readAtendRows(aba, last) {
  const nCols = Math.min(COL.PRODUTO_ESP, aba.getMaxColumns());
  const rows  = aba.getRange(2, 1, last - 1, nCols).getValues();
  if (nCols >= COL.PRODUTO_ESP) return rows;
  return rows.map(r => {
    const out = r.slice();
    while (out.length < COL.PRODUTO_ESP) out.push('');
    return out;
  });
}

// Só faz sentido guardar/exibir o texto livre quando o motivo é o que abre
// o campo — protege contra lixo vindo de chamadas antigas ou fora do fluxo.
function _produtoEspDe(motivo, valor) {
  return String(motivo) === MOTIVO_PRODUTO_ESP ? String(valor || '').trim() : '';
}

// ── MÉTRICAS ──────────────────────────────────────────────
// Regra confirmada com a cliente:
//   • "Atendimentos" conta TUDO, inclusive troca/ajuste puro — a loja
//     atendeu alguém de verdade.
//   • A conversão olha só venda × perda. Troca não é vitória nem derrota,
//     então fica fora do denominador para não diluir a porcentagem.
// Por isso "perdas" é filtro explícito e nunca (total - vendas): com linhas
// de resultado 'troca' na planilha, aquela subtração contaria toda troca
// como perda.
function _taxaConversao(vendas, perdas) {
  const base = vendas + perdas;
  return base > 0 ? Math.round((vendas / base) * 100) : 0;
}

function _metricas(ats) {
  const total  = ats.length;
  const vendas = ats.filter(a => a.resultado === 'venda').length;
  const perdas = ats.filter(a => a.resultado === 'perda').length;
  // Volume de troca/ajuste: conta todo atendimento que envolveu troca,
  // sozinha ou somada a uma venda/perda. Por ser uma marcação sobreposta,
  // não soma com vendas/perdas para dar o total — é outra leitura.
  const trocas = ats.filter(a => a.isTroca).length;
  return { total, vendas, perdas, trocas, conversao: _taxaConversao(vendas, perdas) };
}

function _queryAtendimentos(filtroLojaKey, filtroPeriodo) {
  const aba  = getAba(ABA.ATEND);
  const last = aba.getLastRow();
  if (last <= 1) return [];

  const rows = _readAtendRows(aba, last);

  let limitDate = null;
  if (filtroPeriodo !== 'tudo') {
    limitDate = new Date();
    limitDate.setHours(0, 0, 0, 0);
    if (filtroPeriodo === '7')  limitDate.setDate(limitDate.getDate() - 7);
    if (filtroPeriodo === '30') limitDate.setDate(limitDate.getDate() - 30);
  }

  const result = [];
  rows.forEach(r => {
    const lojaKey  = String(r[COL.LOJA_KEY  - 1]);
    const lojaNome = String(r[COL.LOJA_NOME - 1]);
    const ts       = String(r[COL.TS        - 1]);
    if (filtroLojaKey && filtroLojaKey !== lojaKey) return;
    if (limitDate && new Date(ts) < limitDate) return;
    result.push({
      ts, lojaKey, nomeLoja: lojaNome,
      data:       String(r[COL.DATA        - 1]),
      hora:       String(r[COL.HORA        - 1]),
      vend:       String(r[COL.VEND        - 1]),
      resultado:  String(r[COL.RESULTADO   - 1]),
      motivo:     String(r[COL.MOTIVO      - 1]),
      isTroca:    String(r[COL.IS_TROCA    - 1]) === 'TRUE',
      produtoEsp: String(r[COL.PRODUTO_ESP - 1] || '').trim(),
    });
  });

  return result.reverse();
}

// Lista crua de produtos que os clientes procuraram — sem normalizar, sem
// agrupar e sem cruzar lojas. O vocabulário de cada loja é diferente e o
// time entende os próprios termos; qualquer "limpeza" aqui destruiria
// informação. A loja vai junto justamente por isso.
function _produtosPedidos(ats) {
  return ats
    .filter(a => a.resultado === 'perda' && a.produtoEsp)
    .map(a => ({ loja: a.nomeLoja, produto: a.produtoEsp, ts: a.ts }));
}

// ── FUNÇÕES PÚBLICAS ───────────────────────────────────────
// Daqui para baixo: loja só mexe na própria fila, e sempre com a loja e o
// vendedor resolvidos no servidor (sessao.storeKey / nome), nunca com o que
// chegou do cliente.
function getFila(storeKey) {
  const sessao = _requireLoja(storeKey);
  return { ok: true, fila: _readFila(sessao.storeKey) };
}

function entrarFila(storeKey, vend) {
  const sessao = _requireLoja(storeKey);
  const nome   = _requireVendedorDaLoja(sessao.storeKey, vend);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const fila = _readFila(sessao.storeKey);
    if (fila.find(f => f.vend === nome))
      return { ok: false, msg: nome + ' já está na fila.' };
    fila.push({ vend: nome, status: 'aguardando', entrou: new Date().toISOString() });
    _writeFila(sessao.storeKey, fila);
    SpreadsheetApp.flush();
    return { ok: true, fila };
  } finally { lock.releaseLock(); }
}

function sairFila(storeKey, vend) {
  const sessao = _requireLoja(storeKey);
  const nome   = _requireVendedorDaLoja(sessao.storeKey, vend);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const fila = _readFila(sessao.storeKey).filter(f => f.vend !== nome);
    _writeFila(sessao.storeKey, fila);
    SpreadsheetApp.flush();
    return { ok: true, fila };
  } finally { lock.releaseLock(); }
}

// tipoPausa é opcional: chamadas antigas com 3 argumentos continuam
// válidas. Só tem sentido enquanto status === 'pausa'; em qualquer outro
// status o campo é apagado para a próxima pausa começar limpa.
function setFilaStatus(storeKey, vend, status, tipoPausa) {
  const sessao = _requireLoja(storeKey);
  const nome   = _requireVendedorDaLoja(sessao.storeKey, vend);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const fila = _readFila(sessao.storeKey);
    const item = fila.find(f => f.vend === nome);
    if (item) {
      item.status = status;
      // A posição no array não muda aqui — pausar nunca move ninguém.
      if (status === 'pausa' && PAUSA_TIPOS.indexOf(String(tipoPausa)) !== -1) {
        item.tipoPausa = String(tipoPausa);
      } else {
        delete item.tipoPausa;
      }
    }
    _writeFila(sessao.storeKey, fila);
    SpreadsheetApp.flush();
    return { ok: true, fila };
  } finally { lock.releaseLock(); }
}

// Volta da pausa seguindo as REGRAS de "ausências momentâneas":
//   banheiro  → mantém a posição exata que já tinha (não é movido)
//   pessoal   → vai para o fim da fila
//   intervalo → vai para o fim da fila
// Vale para qualquer posição, não só para quem estava na vez. Sem tipo
// gravado (pausa antiga, anterior a esta mudança) trata como banheiro —
// não mover é a opção segura: nunca faz ninguém furar a fila.
function voltarDaPausa(storeKey, vend) {
  const sessao = _requireLoja(storeKey);
  const nome   = _requireVendedorDaLoja(sessao.storeKey, vend);
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const fila = _readFila(sessao.storeKey);
    const idx  = fila.findIndex(f => f.vend === nome);
    if (idx !== -1) {
      const item = fila[idx];
      const tipo = item.tipoPausa || PAUSA_MANTEM_POSICAO;
      item.status = 'aguardando';
      delete item.tipoPausa;
      if (tipo !== PAUSA_MANTEM_POSICAO) {
        fila.splice(idx, 1);
        fila.push(item);
      }
    }
    _writeFila(sessao.storeKey, fila);
    SpreadsheetApp.flush();
    return { ok: true, fila };
  } finally { lock.releaseLock(); }
}

// produtoEspecifico é opcional (default '') — chamadas antigas com 6
// argumentos continuam válidas e gravam a coluna 11 vazia.
// storeName continua na assinatura só por compatibilidade com o front, mas
// o valor é DESCARTADO de propósito: quem grava o nome da loja no histórico
// é o servidor, senão dava para carimbar o atendimento em outra loja.
function registrarAtendimento(storeKey, storeName, vend, resultado, motivo, isTroca, produtoEspecifico) {
  const sessao = _requireLoja(storeKey);
  const nome   = _requireVendedorDaLoja(sessao.storeKey, vend);
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    _appendAtendimento(sessao.storeKey, sessao.storeName, nome, resultado, motivo, isTroca,
                       _produtoEspDe(motivo, produtoEspecifico));
    const fila = _readFila(sessao.storeKey);
    const idx  = fila.findIndex(f => f.vend === nome);
    if (idx !== -1) {
      const [item] = fila.splice(idx, 1);
      item.status = 'aguardando';
      // A vez só gira quando houve venda de verdade. Troca/ajuste sozinho —
      // ou junto de uma perda — não consome a vez, então o vendedor volta
      // para a frente da fila.
      //   venda            → gira    | perda           → gira
      //   troca            → mantém  | troca + venda   → gira
      //   troca + perda    → mantém
      const rotaciona = (resultado === 'venda') || !isTroca;
      rotaciona ? fila.push(item) : fila.unshift(item);
    }
    _writeFila(sessao.storeKey, fila);
    SpreadsheetApp.flush();
    return { ok: true, fila };
  } finally { lock.releaseLock(); }
}

// ── ANÁLISE INDIVIDUAL DE VENDEDOR ────────────────────────
// Aceita datas livres (yyyy-MM-dd) + filtro de vendedor
function getAnaliseVendedor(vendedor, dataInicio, dataFim) {
  _requireAdmin();   // dado de desempenho individual é só da gestora
  const aba  = getAba(ABA.ATEND);
  const last = aba.getLastRow();
  if (last <= 1) return { ok: true, dados: _analiseVazia(vendedor) };

  const rows = _readAtendRows(aba, last);

  // Monta limite de datas — inclui o dia inteiro do fim
  const dInicio = dataInicio ? new Date(dataInicio + 'T00:00:00') : null;
  const dFim    = dataFim    ? new Date(dataFim    + 'T23:59:59') : null;

  // Filtra registros do vendedor no período
  const ats = [];
  rows.forEach(r => {
    const vend = String(r[COL.VEND - 1]);
    const ts   = String(r[COL.TS   - 1]);
    if (vendedor && vend !== vendedor) return;
    const d = new Date(ts);
    if (dInicio && d < dInicio) return;
    if (dFim    && d > dFim)    return;
    ats.push({
      ts,
      data:       String(r[COL.DATA        - 1]),
      hora:       String(r[COL.HORA        - 1]),
      lojaKey:    String(r[COL.LOJA_KEY    - 1]),
      nomeLoja:   String(r[COL.LOJA_NOME   - 1]),
      vend,
      resultado:  String(r[COL.RESULTADO   - 1]),
      motivo:     String(r[COL.MOTIVO      - 1]),
      isTroca:    String(r[COL.IS_TROCA    - 1]) === 'TRUE',
      produtoEsp: String(r[COL.PRODUTO_ESP - 1] || '').trim(),
    });
  });

  const { total, vendas, perdas, trocas, conversao } = _metricas(ats);

  // Atendimentos por dia (para gráfico)
  const porDiaMap = {};
  ats.forEach(a => {
    if (!porDiaMap[a.data]) porDiaMap[a.data] = { data: a.data, total: 0, vendas: 0, perdas: 0, trocas: 0 };
    porDiaMap[a.data].total++;
    if (a.resultado === 'venda')      porDiaMap[a.data].vendas++;
    else if (a.resultado === 'perda') porDiaMap[a.data].perdas++;
    if (a.isTroca) porDiaMap[a.data].trocas++;
  });
  const porDia = Object.values(porDiaMap).sort((a, b) => a.data.localeCompare(b.data));

  // Motivos de perda
  const motivoMap = {};
  ats.filter(a => a.resultado === 'perda' && a.motivo).forEach(a => {
    motivoMap[a.motivo] = (motivoMap[a.motivo] || 0) + 1;
  });
  const motivos = Object.entries(motivoMap)
    .map(([m, n]) => ({ motivo: m, n }))
    .sort((a, b) => b.n - a.n);

  // Produtos que os clientes procuraram — texto cru, mais recentes primeiro.
  // Lista separada: nunca entra no breakdown de motivos acima.
  const produtosPedidos = _produtosPedidos(ats).reverse();

  // Comparativo: média de conversão de todos os vendedores no mesmo período
  const todosAts = [];
  rows.forEach(r => {
    const ts = String(r[COL.TS - 1]);
    const d  = new Date(ts);
    if (dInicio && d < dInicio) return;
    if (dFim    && d > dFim)    return;
    todosAts.push({
      vend:      String(r[COL.VEND      - 1]),
      resultado: String(r[COL.RESULTADO - 1]),
    });
  });

  const mediaMap = {};
  todosAts.forEach(a => {
    if (!mediaMap[a.vend]) mediaMap[a.vend] = { total: 0, vendas: 0, perdas: 0 };
    mediaMap[a.vend].total++;
    if (a.resultado === 'venda')      mediaMap[a.vend].vendas++;
    else if (a.resultado === 'perda') mediaMap[a.vend].perdas++;
  });
  // Mesma regra da conversão individual: troca fica fora do denominador.
  // Quem só registrou troca no período não entra na média — não tem taxa,
  // e entrar como 0% puxaria a média da equipe para baixo sem motivo.
  const taxas = Object.values(mediaMap)
    .filter(v => (v.vendas + v.perdas) > 0)
    .map(v => (v.vendas / (v.vendas + v.perdas)) * 100);
  const mediaGeral = taxas.length > 0 ? Math.round(taxas.reduce((s, t) => s + t, 0) / taxas.length) : 0;

  // Ranking de todos no período (para comparativo)
  const ranking = Object.entries(mediaMap)
    .map(([v, d]) => ({ vend: v, total: d.total, vendas: d.vendas, tx: _taxaConversao(d.vendas, d.perdas) }))
    .sort((a, b) => b.tx - a.tx);
  const posicao = ranking.findIndex(r => r.vend === vendedor) + 1;

  return {
    ok: true,
    dados: {
      vendedor, total, vendas, perdas, trocas, conversao,
      mediaGeral, posicao, totalVendedores: ranking.length,
      porDia, motivos, produtosPedidos,
      historico: ats.reverse().slice(0, 50),
    }
  };
}

function _analiseVazia(vendedor) {
  return {
    vendedor, total: 0, vendas: 0, perdas: 0, trocas: 0,
    conversao: 0, mediaGeral: 0, posicao: 0, totalVendedores: 0,
    porDia: [], motivos: [], produtosPedidos: [], historico: [],
  };
}

// Retorna vendedores da aba vendedores filtrados por loja
// Só a gestora usa (filtro da Análise individual); a loja já recebe a
// própria equipe pelo verificarAcesso().
function getVendedoresPorLoja(lojaKey) {
  _requireAdmin();
  const aba = getAba(ABA.VEND);
  if (!aba) return { ok: true, vendedores: [] };
  const rows = aba.getDataRange().getValues().slice(1);
  const lista = [];
  rows.forEach(r => {
    const [lk,, vend, ativo] = r;
    if (String(ativo).toUpperCase() !== 'TRUE') return;
    if (lojaKey && String(lk).trim() !== lojaKey) return;
    lista.push(String(vend).trim());
  });
  return { ok: true, vendedores: lista };
}

function getPainelData(filtroLojaKey, filtroPeriodo) {
  _requireAdmin();   // painel consolidado cruza as duas lojas
  const ats = _queryAtendimentos(filtroLojaKey || null, filtroPeriodo || 'hoje');
  const { total, vendas, perdas, trocas, conversao } = _metricas(ats);

  const rankMap = {};
  ats.forEach(a => {
    if (!rankMap[a.vend]) rankMap[a.vend] = { vend: a.vend, loja: a.nomeLoja, total: 0, vendas: 0, perdas: 0 };
    rankMap[a.vend].total++;
    if (a.resultado === 'venda')      rankMap[a.vend].vendas++;
    else if (a.resultado === 'perda') rankMap[a.vend].perdas++;
  });
  const ranking = Object.values(rankMap)
    .map(r => ({ ...r, tx: _taxaConversao(r.vendas, r.perdas) }))
    .sort((a, b) => b.vendas - a.vendas);

  const motivoMap = {};
  ats.filter(a => a.resultado === 'perda' && a.motivo).forEach(a => {
    motivoMap[a.motivo] = (motivoMap[a.motivo] || 0) + 1;
  });
  const motivos = Object.entries(motivoMap)
    .map(([m, n]) => ({ motivo: m, n }))
    .sort((a, b) => b.n - a.n);

  // Produtos que os clientes procuraram — texto cru, já filtrado por
  // loja/período pelo _queryAtendimentos, mais recentes primeiro.
  const produtosPedidos = _produtosPedidos(ats);

  return {
    ok: true,
    total, vendas, perdas, trocas, conversao,
    ranking, motivos, produtosPedidos,
    recentes: ats.slice(0, 30),
    filas: _readTodasFilas(),
  };
}

// ── EXPORTAR PLANILHA ──────────────────────────────────────
function exportarPlanilha(filtroLojaKey, dataInicio, dataFim) {
  try {
    // Dentro do try para o "sem permissão" virar { ok:false, erro } e cair
    // no toast que já existe, em vez de estourar como exceção crua.
    _requireAdmin();

    // ── 1. Query de dados ───────────────────────────────────
    const aba  = getAba(ABA.ATEND);
    const last = aba.getLastRow();
    if (last <= 1) return { ok: false, erro: 'Nenhum registro encontrado.' };

    const rows = _readAtendRows(aba, last);

    const dInicio = dataInicio ? new Date(dataInicio + 'T00:00:00') : null;
    const dFim    = dataFim    ? new Date(dataFim    + 'T23:59:59') : null;

    const ats = [];
    rows.forEach(r => {
      const lojaKey = String(r[COL.LOJA_KEY - 1]);
      const ts      = String(r[COL.TS       - 1]);
      if (filtroLojaKey && filtroLojaKey !== lojaKey) return;
      const d = new Date(ts);
      if (dInicio && d < dInicio) return;
      if (dFim    && d > dFim)    return;
      ats.push({
        data:       String(r[COL.DATA        - 1]),
        hora:       String(r[COL.HORA        - 1]),
        lojaKey,
        nomeLoja:   String(r[COL.LOJA_NOME   - 1]),
        vend:       String(r[COL.VEND        - 1]),
        resultado:  String(r[COL.RESULTADO   - 1]),
        motivo:     String(r[COL.MOTIVO      - 1]),
        isTroca:    String(r[COL.IS_TROCA    - 1]) === 'TRUE',
        produtoEsp: String(r[COL.PRODUTO_ESP - 1] || '').trim(),
      });
    });

    if (!ats.length) return { ok: false, erro: 'Nenhum registro no período selecionado.' };

    // ── 2. Cria a planilha e move para a pasta configurada ──
    const periodoLabel = (dataInicio || 'início') + ' a ' + (dataFim || 'hoje');
    const nome = 'CaliforU — Relatório ' + periodoLabel;

    const ss     = SpreadsheetApp.create(nome);
    const fileId = ss.getId();

    // Move para pasta configurada (só tenta se o ID foi preenchido)
    if (DRIVE_FOLDER_ID && DRIVE_FOLDER_ID !== 'SEU_FOLDER_ID_AQUI') {
      try {
        const file  = DriveApp.getFileById(fileId);
        const pasta = DriveApp.getFolderById(DRIVE_FOLDER_ID);
        pasta.addFile(file);
        DriveApp.getRootFolder().removeFile(file);
      } catch(errDrive) {
        // Se não conseguir mover, mantém no Drive raiz sem quebrar a exportação
        Logger.log('Aviso: não foi possível mover para a pasta. ' + errDrive.message);
      }
    }

    // ── 3. Helpers de estilo ────────────────────────────────
    const HEAD_BG  = '#FF5A36';
    const HEAD_FG  = '#FFFFFF';
    const ALT_BG   = '#FFF4F1';
    const VENDA_BG = '#D4EDDA';
    const VENDA_FG = '#155724';
    const PERDA_BG = '#F8D7DA';
    const PERDA_FG = '#721C24';
    const TROCA_BG = '#E2E3E5';
    const TROCA_FG = '#383D41';

    function cabecalho(sheet, cols) {
      const r = sheet.getRange(1, 1, 1, cols);
      r.setBackground(HEAD_BG)
       .setFontColor(HEAD_FG)
       .setFontWeight('bold')
       .setFontSize(10);
      sheet.setFrozenRows(1);
    }

    function autoResize(sheet, cols) {
      for (let c = 1; c <= cols; c++) sheet.autoResizeColumn(c);
    }

    function colorirResultado(sheet, row, col, resultado) {
      const cell = sheet.getRange(row, col);
      if (resultado === 'venda') {
        cell.setBackground(VENDA_BG).setFontColor(VENDA_FG).setFontWeight('bold');
      } else if (resultado === 'troca') {
        cell.setBackground(TROCA_BG).setFontColor(TROCA_FG).setFontWeight('bold');
      } else {
        cell.setBackground(PERDA_BG).setFontColor(PERDA_FG).setFontWeight('bold');
      }
    }

    // Troca/ajuste puro não é venda nem perda — precisa de rótulo próprio,
    // senão a exportação mostraria "Perda" para quem não perdeu nada.
    function rotuloResultado(resultado) {
      if (resultado === 'venda') return 'Venda';
      if (resultado === 'troca') return 'Troca/Ajuste';
      return 'Perda';
    }

    // ── 4. Aba: Resumo Geral ────────────────────────────────
    const abaResumo = ss.getActiveSheet();
    abaResumo.setName('Resumo Geral');

    const { total, vendas, perdas, trocas, conversao } = _metricas(ats);

    // Bloco de cabeçalho do relatório
    abaResumo.appendRow(['CaliforU — Relatório de Atendimentos', '']);
    abaResumo.getRange(1, 1).setFontSize(14).setFontWeight('bold').setFontColor(HEAD_BG);
    abaResumo.appendRow(['Período', periodoLabel]);
    abaResumo.appendRow(['Gerado em', Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')]);
    abaResumo.appendRow(['Loja', filtroLojaKey || 'Todas']);
    // Linha de separação visual (sem appendRow vazio — usa espaço)
    abaResumo.appendRow([' ', ' ']);

    // Tabela de métricas com cabeçalho na linha 6
    abaResumo.appendRow(['Métrica', 'Valor']);
    abaResumo.getRange(6, 1, 1, 2)
      .setBackground(HEAD_BG).setFontColor(HEAD_FG).setFontWeight('bold');
    abaResumo.appendRow(['Total de atendimentos', total]);
    abaResumo.appendRow(['Vendas realizadas',     vendas]);
    abaResumo.appendRow(['Perdas',                perdas]);
    // Marcação sobreposta: uma troca pode vir junto de uma venda ou perda,
    // por isso esta linha não soma com as de cima para dar o total.
    abaResumo.appendRow(['Atendimentos com troca/ajuste', trocas]);
    abaResumo.appendRow(['Taxa de conversão (venda ÷ venda+perda)', conversao + '%']);
    abaResumo.getRange(abaResumo.getLastRow(), 2).setFontWeight('bold').setFontColor(
      conversao >= 50 ? VENDA_FG : PERDA_FG
    );
    abaResumo.setFrozenRows(0);
    autoResize(abaResumo, 2);

    // ── 5. Abas por loja ────────────────────────────────────
    const lojas = [...new Set(ats.map(a => a.lojaKey))];
    lojas.forEach(lojaKey => {
      const nomeLoja = ats.find(a => a.lojaKey === lojaKey)?.nomeLoja || lojaKey;
      const atLoja   = ats.filter(a => a.lojaKey === lojaKey);
      const sheet    = ss.insertSheet(nomeLoja);

      sheet.appendRow(['Data', 'Hora', 'Vendedor', 'Resultado', 'Motivo de Perda', 'Produto Específico', 'Troca/Ajuste']);
      cabecalho(sheet, 7);

      atLoja.forEach((a, i) => {
        const rowNum = i + 2;
        sheet.appendRow([
          a.data, a.hora, a.vend,
          rotuloResultado(a.resultado),
          a.motivo || '—',
          a.produtoEsp || '—',
          a.isTroca ? 'Sim' : 'Não',
        ]);
        colorirResultado(sheet, rowNum, 4, a.resultado);
        if (i % 2 === 1) {
          sheet.getRange(rowNum, 1, 1, 3).setBackground(ALT_BG);
          sheet.getRange(rowNum, 5, 1, 3).setBackground(ALT_BG);
        }
      });
      autoResize(sheet, 7);
    });

    // ── 6. Aba: Ranking ─────────────────────────────────────
    const abaRank = ss.insertSheet('Ranking');
    abaRank.appendRow(['Posição', 'Vendedor', 'Loja', 'Atendimentos', 'Vendas', 'Perdas', 'Conversão (%)']);
    cabecalho(abaRank, 7);

    const rankMap = {};
    ats.forEach(a => {
      if (!rankMap[a.vend]) rankMap[a.vend] = { vend: a.vend, loja: a.nomeLoja, total: 0, vendas: 0, perdas: 0 };
      rankMap[a.vend].total++;
      if (a.resultado === 'venda')      rankMap[a.vend].vendas++;
      else if (a.resultado === 'perda') rankMap[a.vend].perdas++;
    });

    Object.values(rankMap)
      .map(r => ({ ...r, tx: _taxaConversao(r.vendas, r.perdas) }))
      .sort((a, b) => b.tx - a.tx)
      .forEach((r, i) => {
        const rowNum = i + 2;
        abaRank.appendRow([i + 1, r.vend, r.loja, r.total, r.vendas, r.perdas, r.tx + '%']);
        if (i === 0) abaRank.getRange(rowNum, 1, 1, 7).setBackground('#FFE8B0').setFontWeight('bold');
        else if (i === 1) abaRank.getRange(rowNum, 1, 1, 7).setBackground('#F0F0F0');
        else if (i === 2) abaRank.getRange(rowNum, 1, 1, 7).setBackground('#F5E6D3');
      });
    autoResize(abaRank, 7);

    // ── 7. Aba: Por Vendedor ────────────────────────────────
    const abaPorVend = ss.insertSheet('Por Vendedor');
    abaPorVend.appendRow(['Vendedor', 'Loja', 'Data', 'Hora', 'Resultado', 'Motivo', 'Produto Específico']);
    cabecalho(abaPorVend, 7);

    let rowAtual = 2;
    [...new Set(ats.map(a => a.vend))].sort().forEach(vend => {
      const atVend = ats.filter(a => a.vend === vend);
      atVend.forEach((a, i) => {
        abaPorVend.appendRow([
          i === 0 ? a.vend    : '',
          i === 0 ? a.nomeLoja : '',
          a.data, a.hora,
          rotuloResultado(a.resultado),
          a.motivo || '—',
          a.produtoEsp || '—',
        ]);
        colorirResultado(abaPorVend, rowAtual, 5, a.resultado);
        rowAtual++;
      });
      // Separador entre vendedores — linha com fundo cinza, sem appendRow vazio
      abaPorVend.appendRow([' ', ' ', ' ', ' ', ' ', ' ', ' ']);
      abaPorVend.getRange(rowAtual, 1, 1, 7).setBackground('#EEEEEE');
      rowAtual++;
    });
    autoResize(abaPorVend, 7);

    // ── 8. Aba: Motivos de Perda ─────────────────────────────
    const abaMotivos = ss.insertSheet('Motivos de Perda');
    abaMotivos.appendRow(['Motivo', 'Qtd Total', 'Loja', 'Vendedor', 'Qtd por Vendedor']);
    cabecalho(abaMotivos, 5);

    const motivoGeral = {};
    ats.filter(a => a.resultado === 'perda' && a.motivo).forEach(a => {
      if (!motivoGeral[a.motivo]) motivoGeral[a.motivo] = { total: 0, porVend: {} };
      motivoGeral[a.motivo].total++;
      const kv = a.vend + '||' + a.nomeLoja;
      motivoGeral[a.motivo].porVend[kv] = (motivoGeral[a.motivo].porVend[kv] || 0) + 1;
    });

    Object.entries(motivoGeral)
      .sort((a, b) => b[1].total - a[1].total)
      .forEach(([motivo, dados]) => {
        let primeiro = true;
        Object.entries(dados.porVend)
          .sort((a, b) => b[1] - a[1])
          .forEach(([chave, qtd]) => {
            const [vend, loja] = chave.split('||');
            abaMotivos.appendRow([
              primeiro ? motivo    : '',
              primeiro ? dados.total : '',
              loja, vend, qtd,
            ]);
            primeiro = false;
          });
      });
    autoResize(abaMotivos, 5);

    // ── 9. Retorna link ─────────────────────────────────────
    SpreadsheetApp.flush();
    return { ok: true, url: ss.getUrl(), nome: ss.getName() };

  } catch(err) {
    return { ok: false, erro: err.message };
  }
}

// ── CONFIGURAÇÕES: GERENCIAR VENDEDORES ───────────────────

// Retorna todos os vendedores agrupados por loja para a tela de config
function getConfigVendedores() {
  _requireAdmin();   // tela de configurações é exclusiva da gestora
  const aba = getAba(ABA.VEND);
  if (!aba) return { ok: false, erro: 'Aba vendedores não encontrada.' };

  const rows = aba.getDataRange().getValues().slice(1);
  const lojas = {};

  // Inicializa APENAS com as lojas da whitelist — sem duplicar
  _getLojas().forEach(l => {
    lojas[l.key] = { nome: l.name, vendedores: [] };
  });

  // Adiciona vendedores apenas nas lojas que existem na whitelist
  rows.forEach((r, i) => {
    const [lojaKey,, vend, ativo] = r;
    const k = String(lojaKey).trim();
    if (!k || !lojas[k]) return; // ignora lojas desconhecidas (linhas antigas)
    if (String(ativo).toUpperCase() !== 'TRUE') return;
    lojas[k].vendedores.push({
      nome: String(vend).trim(),
      ativo: true,
      row: i + 2,
    });
  });

  return { ok: true, lojas };
}

// Adiciona um novo vendedor à aba vendedores
function adicionarVendedor(lojaKey, nomeVend) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    // Dentro do try para o "sem permissão" sair como { ok:false, erro },
    // que é o formato que a tela de configurações já sabe mostrar.
    _requireAdmin();

    const aba      = getAba(ABA.VEND);
    const lojaInfo = _getLojas().find(l => l.key === lojaKey);
    if (!lojaInfo) return { ok: false, erro: 'Loja não encontrada.' };

    const nome = String(nomeVend).trim();
    if (!nome) return { ok: false, erro: 'Nome inválido.' };

    // Verifica duplicata na mesma loja
    const rows = aba.getDataRange().getValues().slice(1);
    const existe = rows.some(r =>
      String(r[0]).trim() === lojaKey &&
      String(r[2]).trim().toLowerCase() === nome.toLowerCase() &&
      String(r[3]).toUpperCase() === 'TRUE'
    );
    if (existe) return { ok: false, erro: nome + ' já está cadastrado nesta loja.' };

    aba.appendRow([lojaKey, lojaInfo.name, nome, 'TRUE']);
    SpreadsheetApp.flush();
    return { ok: true };
  } catch(err) {
    return { ok: false, erro: err.message };
  } finally {
    lock.releaseLock();
  }
}

// Remove vendedor: desativa na aba vendedores + apaga histórico de atendimentos
function removerVendedor(lojaKey, nomeVend) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    // Idem: erro de permissão vira { ok:false, erro }. Aqui importa ainda
    // mais — esta função apaga histórico e é irreversível.
    _requireAdmin();

    const nome = String(nomeVend).trim();

    // 1. Desativa na aba vendedores (marca como FALSE)
    const abaVend = getAba(ABA.VEND);
    const lastVend = abaVend.getLastRow();
    if (lastVend >= 2) {
      const rows = abaVend.getRange(2, 1, lastVend - 1, 4).getValues();
      for (let i = rows.length - 1; i >= 0; i--) {
        if (String(rows[i][0]).trim() === lojaKey &&
            String(rows[i][2]).trim().toLowerCase() === nome.toLowerCase()) {
          // Deleta a linha inteira
          abaVend.deleteRow(i + 2);
        }
      }
    }

    // 2. Apaga histórico de atendimentos do vendedor
    const abaAt  = getAba(ABA.ATEND);
    const lastAt = abaAt.getLastRow();
    if (lastAt >= 2) {
      const colVend = abaAt.getRange(2, COL.VEND, lastAt - 1, 1).getValues();
      // Deleta de baixo pra cima para não deslocar índices
      for (let i = colVend.length - 1; i >= 0; i--) {
        if (String(colVend[i][0]).trim().toLowerCase() === nome.toLowerCase()) {
          abaAt.deleteRow(i + 2);
        }
      }
    }

    // 3. Remove da fila se estiver ativo
    const fila = _readFila(lojaKey).filter(f => f.vend !== nome);
    _writeFila(lojaKey, fila);

    SpreadsheetApp.flush();
    return { ok: true };
  } catch(err) {
    return { ok: false, erro: err.message };
  } finally {
    lock.releaseLock();
  }
}
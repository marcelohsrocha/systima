/**
 * ============================================================
 *  QUADRO DE DEMANDAS — Google Forms → Kanban + Google Agenda
 *  Template replicável · Studio Ponte Marketing e Design
 * ============================================================
 *  Tudo que é específico de cada operação fica nas abas:
 *  Config · Categorias · Equipe · Mapeamento · Automacoes · Recorrentes
 *  O código não precisa ser alterado para replicar.
 * ============================================================
 */

/** ID da planilha. Necessário quando o script é um projeto avulso (não vinculado à planilha). */
const PLANILHA_ID = ''; // ex.: '1AbC...xyz' (trecho entre /d/ e /edit na URL da planilha). Deixe vazio se o script for vinculado à planilha.

let _ss = null;
function ss_() {
  if (_ss) return _ss;
  try { _ss = SpreadsheetApp.getActive(); } catch (e) { _ss = null; }
  if (!_ss) {
    if (!PLANILHA_ID) throw new Error('Defina PLANILHA_ID no topo do Code.gs (projeto avulso) ou crie o script pela planilha (Extensões → Apps Script).');
    _ss = SpreadsheetApp.openById(PLANILHA_ID);
  }
  return _ss;
}

/** Mostra aviso na planilha quando houver interface; senão, registra no log de execução. */
function alerta_(msg) {
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
  return msg;
}

/** Bloqueia funções administrativas quando chamadas por terceiros pelo app da web. */
function soDono_() {
  const ativo = Session.getActiveUser().getEmail();
  const dono = Session.getEffectiveUser().getEmail();
  if (!ativo || ativo !== dono) throw new Error('Função administrativa: rode pelo editor do Apps Script com a conta dona da planilha.');
}

const ABA = {
  CONFIG: 'Config',
  CATEGORIAS: 'Categorias',
  EQUIPE: 'Equipe',
  MAPEAMENTO: 'Mapeamento',
  AUTOMACOES: 'Automacoes',
  RECORRENTES: 'Recorrentes',
  TAREFAS: 'Tarefas',
  COMENTARIOS: 'Comentarios',
  HISTORICO: 'Historico',
  DESDOBRAMENTO: 'Desdobramento'
};

// [chave interna, cabeçalho na planilha]
const CAMPOS = [
  ['id', 'ID'],
  ['criadoEm', 'Criado em'],
  ['titulo', 'Título'],
  ['categoria', 'Categoria'],
  ['status', 'Status'],
  ['responsavel', 'Responsável'],
  ['inicio', 'Início'],
  ['prazo', 'Prazo'],
  ['prioridade', 'Prioridade'],
  ['etiquetas', 'Etiquetas'],
  ['solicitante', 'Solicitante'],
  ['contato', 'Contato'],
  ['emailSolicitante', 'E-mail solicitante'],
  ['briefing', 'Briefing (JSON)'],
  ['anexos', 'Anexos (JSON)'],
  ['checklist', 'Checklist (JSON)'],
  ['estimativa', 'Estimativa (h)'],
  ['tempoMin', 'Tempo gasto (min)'],
  ['timerInicio', 'Timer início'],
  ['timerAutor', 'Timer autor'],
  ['eventoId', 'Evento Agenda ID'],
  ['arquivada', 'Arquivada'],
  ['concluidaEm', 'Concluída em'],
  ['atualizadoEm', 'Atualizado em'],
  ['ordem', 'Ordem'],
  ['origem', 'Origem']
];
const CAMPOS_JSON = ['briefing', 'anexos', 'checklist'];

const CONFIG_PADRAO = [
  ['NOME_QUADRO', 'Sýstima', 'Nome exibido no topo do quadro e nos e-mails'],
  ['COLUNAS', 'Novo | Em produção | Revisão | Entregue', 'Colunas do Kanban, separadas por |'],
  ['COLUNA_FINAL', 'Entregue', 'Coluna que significa "concluído" (não gera alerta de atraso)'],
  ['PRIORIDADES', 'Baixa | Normal | Alta | Urgente', 'Níveis de prioridade, separados por |'],
  ['PRIORIDADE_PADRAO', 'Normal', ''],
  ['COR_PRINCIPAL', '#1F3A5F', 'Cor da marca no quadro (hex)'],
  ['EMAIL_COORDENACAO', '', 'E-mail do gestor do setor (recebe resumo de novas demandas e atrasos)'],
  ['NOTIFICAR_RESPONSAVEL', 'SIM', 'E-mail ao responsável quando recebe uma demanda'],
  ['NOTIFICAR_SOLICITANTE', 'SIM', 'E-mail ao solicitante quando o status muda (se o formulário coletar e-mail)'],
  ['WEBHOOK_NOTIFICACAO', '', 'URL de webhook (Make/Z-API) para enviar WhatsApp. Deixe vazio para desativar'],
  ['AGENDA_INTEGRAR', 'SIM', 'Criar eventos de entrega no Google Agenda'],
  ['AGENDA_ID', '', 'ID da agenda do setor. Vazio = agenda principal da conta do gestor'],
  ['AGENDA_HORA_ENTREGA', '', 'Ex.: 18:00 cria evento com horário. Vazio = evento de dia inteiro'],
  ['AGENDA_CONVIDAR_RESPONSAVEL', 'SIM', 'Adiciona o responsável como convidado (aparece na agenda pessoal dele)'],
  ['AGENDA_CONVITE_POR_EMAIL', 'NAO', 'Enviar também o e-mail de convite do Google Agenda'],
  ['ARQUIVAR_APOS_DIAS', '14', 'Arquiva automaticamente tarefas concluídas há X dias (0 = nunca)'],
  ['EXIGIR_LOGIN', 'SIM', 'Cada pessoa entra com nome + PIN (aba Equipe). Necessário para os perfis funcionarem'],
  ['CODIGO_ACESSO', '', 'Opcional: código geral exigido para abrir o quadro'],
  ['AGENDA_EXIBIR', 'principal', 'Agendas mostradas no Calendário do quadro (IDs separados por vírgula; principal = agenda da conta dona)'],
  ['ABA_RESPOSTAS', '', 'Nome da aba de respostas do Forms. Vazio = detecta automaticamente']
];

const CATEGORIAS_PADRAO = [
  ['Design', '', '#7C3AED', 5, 'Receber briefing; Criar layout; Revisão interna; Aprovação do solicitante; Exportar e entregar'],
  ['Vídeo', '', '#DC2626', 7, 'Roteiro; Captação; Edição; Revisão; Entrega/publicação'],
  ['Fotografia', '', '#0891B2', 3, 'Escalar fotógrafo; Cobertura; Seleção e edição; Entrega do álbum'],
  ['Redes sociais', '', '#DB2777', 3, 'Definir formato; Produzir conteúdo; Legenda; Aprovação; Publicar'],
  ['Suporte técnico', '', '#475467', 5, 'Entender a solicitação; Escalar equipe; Executar; Validar com o solicitante']
];

/* Preencha nomes e e-mails; o PIN é gerado automaticamente. Perfis: Gestor | Membro | Leitor */
const EQUIPE_PADRAO = [
  ['Gestor', '', 'Coordenação', '', 'Gestor', '']
];

const MAPEAMENTO_PADRAO = [
  ['titulo', '', 'Ex.: "Nome do evento / demanda"'],
  ['categoria', '', 'Ex.: "Tipo de demanda" (opções iguais às da aba Categorias)'],
  ['prazo', '', 'Ex.: "Data de entrega desejada"'],
  ['solicitante', '', 'Ex.: "Seu nome"'],
  ['solicitanteComplemento', '', 'Opcional: aparece entre parênteses ao lado do solicitante (ex.: "Setor")'],
  ['contato', '', 'Ex.: "WhatsApp"'],
  ['emailSolicitante', '', 'Ex.: "E-mail" (ou ative a coleta de e-mail no Forms)'],
  ['prioridade', '', 'Opcional'],
  ['responsavel', '', 'Opcional — se o form já indicar o responsável']
];

/* Uma resposta do Forms pode gerar vários cards (um por área). Exemplos desativados (NAO) — adapte às perguntas do seu formulário.
   [Ativa, Pergunta (contém), Se a resposta contém (* = qualquer), Categoria, Prazo pela pergunta, Dias antes] */
const DESDOBRAMENTO_PADRAO = [
  ['NAO', 'Materiais necessários', 'Arte', 'Design', 'Data do evento', 3],
  ['NAO', 'Materiais necessários', 'Vídeo', 'Vídeo', 'Data do evento', 3],
  ['NAO', 'Materiais necessários', 'Fotografia', 'Fotografia', 'Data do evento', 0],
  ['NAO', 'Materiais necessários', 'Post', 'Redes sociais', 'Data do evento', 2]
];

const AUTOMACOES_PADRAO = [
  ['SIM', 'Status mudou para', 'Revisão', '', '', 'Notificar e-mail', 'coordenacao'],
  ['SIM', 'Tarefa criada', '', 'Prioridade', 'Urgente', 'Notificar e-mail', 'coordenacao'],
  ['SIM', 'Prazo vencido', '', '', '', 'Definir prioridade', 'Alta'],
  ['NAO', 'Tarefa criada', '', 'Categoria', 'Vídeo', 'Adicionar etiqueta', 'audiovisual']
];

const RECORRENTES_PADRAO = [
  ['NAO', 'Relatório semanal de redes', 'Redes sociais', '', 'Semanal', 'seg', 2, 'Métricas da semana anterior', ''],
  ['NAO', 'Planejamento de conteúdo do mês', 'Redes sociais', '', 'Mensal', 25, 5, 'Calendário editorial do próximo mês', '']
];

/* ============================================================
 *  MENU E INSTALAÇÃO
 * ============================================================ */

function onOpen() {
  try { SpreadsheetApp.getUi(); } catch (e) { return; }
  SpreadsheetApp.getUi().createMenu('📋 Quadro')
    .addItem('1. Configurar planilha', 'configurarPlanilha')
    .addItem('2. Ativar automações (gatilhos)', 'ativarGatilhos')
    .addSeparator()
    .addItem('Importar respostas já existentes', 'importarRespostasExistentes')
    .addItem('Importar do Trello', 'abrirImportacaoTrello')
    .addItem('Gerar PINs que faltam', 'gerarPinsMenu')
    .addItem('Enviar acesso (link + PIN) para a equipe', 'enviarAcessos')
    .addItem('Sincronizar tudo com o Google Agenda', 'sincronizarTodasAgenda')
    .addItem('Rodar rotina diária agora', 'rotinaDiaria')
    .addItem('Abrir quadro', 'abrirQuadro')
    .addToUi();
}

function configurarPlanilha() {
  soDono_();
  const ss = ss_();
  _ss = ss;
  criarAba_(ss, ABA.CONFIG, ['Chave', 'Valor', 'Descrição'], CONFIG_PADRAO);
  criarAba_(ss, ABA.CATEGORIAS, ['Categoria', 'Responsável padrão', 'Cor', 'SLA (dias)', 'Checklist padrão (itens separados por ;)'], CATEGORIAS_PADRAO);
  criarAba_(ss, ABA.EQUIPE, ['Nome', 'E-mail', 'Função', 'Capacidade semanal (h)', 'Perfil (Gestor/Membro/Leitor)', 'PIN'], EQUIPE_PADRAO);
  gerarPins_();
  criarAba_(ss, ABA.MAPEAMENTO, ['Campo do sistema', 'Pergunta do formulário (texto exato)', 'Dica'], MAPEAMENTO_PADRAO);
  criarAba_(ss, ABA.DESDOBRAMENTO, ['Ativa', 'Pergunta do formulário (contém)', 'Se a resposta contém (* = qualquer)', 'Gera card na categoria', 'Prazo pela pergunta', 'Dias antes'], DESDOBRAMENTO_PADRAO);
  criarAba_(ss, ABA.AUTOMACOES, ['Ativa', 'Quando', 'Valor do gatilho', 'Se campo', 'For igual a', 'Então', 'Valor da ação'], AUTOMACOES_PADRAO);
  criarAba_(ss, ABA.RECORRENTES, ['Ativa', 'Título', 'Categoria', 'Responsável', 'Frequência', 'Dia', 'Prazo (dias)', 'Descrição', 'Última criação'], RECORRENTES_PADRAO);
  const t = criarAba_(ss, ABA.TAREFAS, CAMPOS.map(c => c[1]), []);
  t.getRange(1, 1, t.getMaxRows(), CAMPOS.length).setNumberFormat('@');
  criarAba_(ss, ABA.COMENTARIOS, ['ID', 'Tarefa ID', 'Autor', 'Texto', 'Data'], []);
  criarAba_(ss, ABA.HISTORICO, ['Data', 'Tarefa ID', 'Autor', 'Ação', 'Detalhe'], []);
  validacoes_(ss);
  alerta_('✅ Planilha configurada.\n\nPreencha as abas Equipe, Categorias e Mapeamento. Depois use "2. Ativar automações".');
}

function criarAba_(ss, nome, cabecalho, linhas) {
  let sh = ss.getSheetByName(nome);
  if (sh) return sh;
  sh = ss.insertSheet(nome);
  sh.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho])
    .setFontWeight('bold').setBackground('#1F3A5F').setFontColor('#FFFFFF');
  sh.setFrozenRows(1);
  if (linhas.length) sh.getRange(2, 1, linhas.length, cabecalho.length).setValues(linhas);
  sh.autoResizeColumns(1, cabecalho.length);
  return sh;
}

function validacoes_(ss) {
  const simNao = SpreadsheetApp.newDataValidation().requireValueInList(['SIM', 'NAO']).build();
  const auto = ss.getSheetByName(ABA.AUTOMACOES);
  auto.getRange('A2:A100').setDataValidation(simNao);
  auto.getRange('B2:B100').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Tarefa criada', 'Status mudou para', 'Responsável mudou', 'Prazo vencido']).build());
  auto.getRange('D2:D100').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Categoria', 'Prioridade', 'Responsável', 'Status', 'Etiqueta']).setAllowInvalid(true).build());
  auto.getRange('F2:F100').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Mover para', 'Atribuir a', 'Definir prioridade', 'Adicionar etiqueta', 'Adicionar item de checklist', 'Notificar e-mail', 'Enviar webhook']).build());
  const eq = ss.getSheetByName(ABA.EQUIPE);
  eq.getRange('E2:E200').setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList(['Gestor', 'Membro', 'Leitor']).build());
  eq.getRange('F2:F200').setNumberFormat('@');
  const rec = ss.getSheetByName(ABA.RECORRENTES);
  rec.getRange('A2:A100').setDataValidation(simNao);
  rec.getRange('E2:E100').setDataValidation(SpreadsheetApp.newDataValidation()
    .requireValueInList(['Diária', 'Semanal', 'Mensal']).build());
  const map = ss.getSheetByName(ABA.MAPEAMENTO);
  const aba = abaRespostas_();
  if (aba && aba.getLastColumn() > 0) {
    const perguntas = aba.getRange(1, 1, 1, aba.getLastColumn()).getDisplayValues()[0].filter(String);
    map.getRange('B2:B20').setDataValidation(SpreadsheetApp.newDataValidation()
      .requireValueInList(perguntas).setAllowInvalid(true).build());
  }
}

function ativarGatilhos() {
  soDono_();
  const ss = ss_();
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger('aoEnviarFormulario').forSpreadsheet(ss).onFormSubmit().create();
  ScriptApp.newTrigger('rotinaDiaria').timeBased().everyDays(1).atHour(7).create();
  validacoes_(ss);
  alerta_('✅ Automações ativas:\n• Nova resposta do Forms → card no quadro\n• Rotina diária às 7h (atrasos, recorrentes, arquivamento)');
}

function abrirQuadro() {
  const url = ScriptApp.getService().getUrl();
  const html = url
    ? `<p>Abra o quadro:</p><p><a href="${url}" target="_blank">${url}</a></p>`
    : '<p>O quadro ainda não foi implantado. Vá em <b>Implantar → Nova implantação → App da Web</b>.</p>';
  try { SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(420).setHeight(140), 'Quadro de Demandas'); }
  catch (e) { console.log(url || 'Quadro ainda não implantado'); }
}

/* ============================================================
 *  WEB APP
 * ============================================================ */

function doGet(e) {
  const pagina = e && e.parameter && e.parameter.pagina === 'importar' ? 'Importar' : 'Index';
  return HtmlService.createTemplateFromFile(pagina).evaluate()
    .setTitle(cfg_().NOME_QUADRO || 'Quadro de Demandas')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** Canal HTTP usado pelo quadro: não envia as contas Google do navegador (evita o bug de várias contas logadas). */
function doPost(e) {
  let saida;
  try {
    const d = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    saida = { ok: true, r: api(d.acao, d.p || {}, d.codigo || '') };
  } catch (err) {
    saida = { ok: false, erro: String((err && err.message) || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(saida)).setMimeType(ContentService.MimeType.JSON);
}

/** Ponto único de entrada do front-end. */
function api(acao, p, codigo) {
  const c = cfg_();
  if (c.CODIGO_ACESSO && String(codigo || '') !== String(c.CODIGO_ACESSO)) throw new Error('ACESSO_NEGADO');
  p = p || {};
  if (acao === 'equipeLogin') return { nome: c.NOME_QUADRO, cor: c.COR_PRINCIPAL, pessoas: equipe_().map(m => ({ nome: m.nome, funcao: m.funcao })) };
  const me = autenticar_(p.autor, p.pin);
  const exigir = c.EXIGIR_LOGIN === 'SIM';
  if (exigir && !me) throw new Error('LOGIN');
  if (acao === 'board') { const b = board_(p); b.me = me ? { nome: me.nome, perfil: me.perfil, interinoAte: me.interinoAte || '' } : null; return b; }
  if (acao === 'detalhes') return detalhes_(p.id);
  if (acao === 'agenda') return agendaEventos_(p.ini, p.fim);

  // daqui pra baixo: escrita
  const eu = me || (exigir ? null : { nome: p.autor || 'Alguém', perfil: 'Gestor' });
  if (!eu) throw new Error('LOGIN');
  if (eu.perfil === 'Leitor') throw new Error('SEM_PERMISSAO: seu perfil é de leitura');
  p.autor = eu.nome;
  if (acao === 'limparQuadro') {
    if (eu.perfil !== 'Gestor') throw new Error('SEM_PERMISSAO: só gestores podem limpar o quadro');
    return limparQuadro_(p, eu);
  }
  if (acao.indexOf('equipe') === 0) {
    if (eu.perfil !== 'Gestor') throw new Error('SEM_PERMISSAO: só gestores gerenciam a equipe');
    return equipeAcao_(acao, p, eu);
  }
  if (acao.indexOf('trello') === 0) {
    if (eu.perfil !== 'Gestor') throw new Error('SEM_PERMISSAO: só gestores importam do Trello');
    if (acao === 'trelloContexto') return trelloContexto_();
    if (acao === 'trelloBuscarApi') return trelloBuscarApi_(p.key, p.token, p.link);
    if (acao === 'trelloImportar') return trelloImportar_(p.dados);
  }
  if (['mover', 'atualizar', 'timer', 'tempoManual'].indexOf(acao) >= 0) checarPermissao_(eu, p.id, acao === 'atualizar' ? (p.campos || {}) : null);

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    switch (acao) {
      case 'mover': return toClient_(alterar_(p.id, { status: p.status, ordem: p.ordem }, p.autor));
      case 'atualizar': return toClient_(alterar_(p.id, p.campos || {}, p.autor));
      case 'criar':
        return toClient_(criarTarefa_(p.dados || {}, p.autor, 'manual'));
      case 'comentar': return comentar_(p.id, p.autor, p.texto);
      case 'timer': return toClient_(timer_(p.id, p.modo, p.autor));
      case 'tempoManual': return toClient_(tempoManual_(p.id, p.minutos, p.autor));
      default: throw new Error('Ação desconhecida: ' + acao);
    }
  } finally {
    lock.releaseLock();
  }
}

/* ---------- perfis e login ---------- */

function autenticar_(nome, pin) {
  const m = membro_(nome);
  if (!m) return null;
  if (m.pin && String(pin || '').trim() !== String(m.pin).trim()) return null;
  return m;
}

/** Gestor: tudo. Membro: altera as próprias demandas (ou sem responsável), pode repassá-las a qualquer pessoa, não arquiva. */
function checarPermissao_(eu, id, campos) {
  if (eu.perfil === 'Gestor') return;
  const t = buscarTarefa_(id);
  if (!t) throw new Error('Tarefa não encontrada');
  if (t.responsavel && norm_(t.responsavel) !== norm_(eu.nome)) throw new Error('SEM_PERMISSAO: você só pode alterar as suas demandas');
  if (campos) {
    if ('arquivada' in campos) throw new Error('SEM_PERMISSAO: só gestores arquivam demandas');
  }
}

function gerarPins_() {
  const sh = ss_().getSheetByName(ABA.EQUIPE);
  if (!sh || sh.getLastRow() < 2) return 0;
  const rg = sh.getRange(2, 1, sh.getLastRow() - 1, 6);
  const v = rg.getDisplayValues();
  let n = 0;
  v.forEach(r => { if (r[0] && !r[5]) { r[5] = String(Math.floor(1000 + Math.random() * 9000)); n++; } if (r[0] && !r[4]) r[4] = 'Membro'; });
  sh.getRange(2, 5, v.length, 2).setNumberFormat('@').setValues(v.map(r => [r[4], r[5]]));
  return n;
}

function gerarPinsMenu() {
  soDono_();
  const n = gerarPins_();
  alerta_(`✅ ${n} PIN(s) gerado(s). Use "Enviar acesso" para mandar por e-mail.`);
}

function enviarAcessos() {
  soDono_();
  const url = ScriptApp.getService().getUrl();
  if (!url) return alerta_('Implante o app da web primeiro (Implantar → Nova implantação → App da Web).');
  gerarPins_();
  const pessoas = equipe_().filter(m => m.email);
  const c = cfg_();
  pessoas.forEach(m => enviarEmail_(m.email, `🔑 Seu acesso ao ${c.NOME_QUADRO}`,
    `<p>Olá, ${esc_(m.nome)}!</p><p>Este é o quadro de demandas da equipe. Lá você acompanha suas entregas, comenta e atualiza o status.</p>
     <p><b>Seu nome:</b> ${esc_(m.nome)}<br><b>Seu PIN:</b> <span style="font-size:20px;letter-spacing:3px">${esc_(m.pin)}</span><br><b>Perfil:</b> ${esc_(m.perfil)}</p>
     <p><a href="${url}" style="background:#1F3A5F;color:#fff;padding:10px 16px;border-radius:6px;text-decoration:none">Abrir o quadro</a></p>
     <p style="color:#6b7280;font-size:13px">No celular, abra o link e use "Adicionar à tela inicial". Não compartilhe seu PIN.</p>`));
  alerta_(`✅ Acesso enviado para ${pessoas.length} pessoa(s): ${pessoas.map(m => m.nome).join(', ')}`);
}

/* ---------- gestão da equipe pelo quadro (Gestor) ---------- */

function equipeAcao_(acao, p, eu) {
  const sh = ss_().getSheetByName(ABA.EQUIPE);
  if (!sh.getRange(1, 7).getDisplayValue()) sh.getRange(1, 7).setValue('Ativo').setFontWeight('bold').setBackground('#1F3A5F').setFontColor('#FFFFFF');
  if (!sh.getRange(1, 8).getDisplayValue()) sh.getRange(1, 8).setValue('Gestor interino até').setFontWeight('bold').setBackground('#1F3A5F').setFontColor('#FFFFFF');
  const linhas = () => sh.getLastRow() < 2 ? [] : sh.getRange(2, 1, sh.getLastRow() - 1, 8).getDisplayValues();
  const achar = nome => linhas().findIndex(r => norm_(r[0]) === norm_(nome));
  const ativo = r => String(r[6] || 'SIM').toUpperCase() !== 'NAO';
  const gestores = () => linhas().filter(r => r[0] && ativo(r) && r[4] === 'Gestor').length;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    let info = '';
    if (acao === 'equipeSalvar') {
      const d = p.dados || {};
      const nome = String(d.nome || '').trim();
      const perfil = ['Gestor', 'Membro', 'Leitor'].indexOf(d.perfil) >= 0 ? d.perfil : 'Membro';
      const email = String(d.email || '').trim();
      if (!nome) throw new Error('Informe o nome');
      if (email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email)) throw new Error('E-mail inválido');
      if (p.original) {
        const i = achar(p.original);
        if (i < 0) throw new Error('Pessoa não encontrada');
        const r = linhas()[i];
        if (norm_(r[0]) === norm_(eu.nome) && perfil !== 'Gestor') throw new Error('Você não pode tirar o seu próprio perfil de Gestor');
        if (r[4] === 'Gestor' && perfil !== 'Gestor' && gestores() <= 1) throw new Error('A equipe precisa de pelo menos um Gestor');
        if (eu.interinoAte && perfil === 'Gestor' && r[4] !== 'Gestor') throw new Error('SEM_PERMISSAO: gestor interino não promove ninguém a Gestor permanente');
        sh.getRange(i + 2, 2, 1, 4).setNumberFormat('@').setValues([[email, String(d.funcao || ''), String(d.capacidade || ''), perfil]]);
        info = 'Alterações salvas.';
      } else {
        if (achar(nome) >= 0) throw new Error('Já existe alguém com esse nome. Use um nome diferente (ex.: "Ana S.").');
        if (eu.interinoAte && perfil === 'Gestor') throw new Error('SEM_PERMISSAO: gestor interino não cadastra novos Gestores');
        const pin = gerarPin_();
        sh.getRange(sh.getLastRow() + 1, 1, 1, 7).setNumberFormat('@')
          .setValues([[nome, email, String(d.funcao || ''), String(d.capacidade || ''), perfil, pin, 'SIM']]);
        info = `${nome} adicionado(a). PIN: ${pin}`;
      }
    } else if (acao === 'equipeNovoPin') {
      const i = achar(p.nome);
      if (i < 0) throw new Error('Pessoa não encontrada');
      const pin = gerarPin_();
      sh.getRange(i + 2, 6).setNumberFormat('@').setValue(pin);
      info = `Novo PIN de ${p.nome}: ${pin}`;
    } else if (acao === 'equipeEnviarAcesso') {
      const m = equipe_().find(x => norm_(x.nome) === norm_(p.nome));
      if (!m || !m.email) throw new Error('Cadastre o e-mail da pessoa antes de enviar o acesso');
      enviarAcessoPara_(m);
      info = `Acesso enviado para ${m.email}.`;
    } else if (acao === 'equipeDelegar') {
      if (eu.interinoAte) throw new Error('SEM_PERMISSAO: só o Gestor titular delega a gestão');
      const i = achar(p.nome);
      if (i < 0) throw new Error('Pessoa não encontrada');
      const r = linhas()[i];
      const ate = String(p.ate || '');
      if (!/^\d{4}-\d{2}-\d{2}$/.test(ate) || ate < hoje_()) throw new Error('Escolha uma data a partir de hoje');
      if (r[4] === 'Gestor') throw new Error(`${r[0]} já é Gestor`);
      if (!ativo(r)) throw new Error('Reative a pessoa antes de delegar');
      sh.getRange(i + 2, 8).setNumberFormat('@').setValue(ate);
      if (r[1]) enviarEmail_(r[1], `🛡️ Você é Gestor interino do ${cfg_().NOME_QUADRO} até ${dataBr_(ate)}`,
        `<p>Olá, ${esc_(r[0])}!</p><p><b>${esc_(eu.nome)}</b> delegou a gestão do quadro para você até <b>${dataBr_(ate)}</b>.</p>
         <p>Nesse período você pode atribuir e reatribuir demandas, arquivar, gerenciar a equipe e recebe os avisos da coordenação. Depois da data, seu perfil volta ao normal automaticamente.</p>
         ${ScriptApp.getService().getUrl() ? `<p><a href="${ScriptApp.getService().getUrl()}">Abrir o quadro</a></p>` : ''}`);
      info = `${r[0]} é Gestor interino até ${dataBr_(ate)}.`;
    } else if (acao === 'equipeEncerrarDelegacao') {
      if (eu.interinoAte && norm_(p.nome) !== norm_(eu.nome)) throw new Error('SEM_PERMISSAO: só o Gestor titular encerra delegações');
      const i = achar(p.nome);
      if (i < 0) throw new Error('Pessoa não encontrada');
      sh.getRange(i + 2, 8).setValue('');
      info = `Delegação de ${linhas()[i][0]} encerrada.`;
    } else if (acao === 'equipeAtivo') {
      const i = achar(p.nome);
      if (i < 0) throw new Error('Pessoa não encontrada');
      const r = linhas()[i];
      if (!p.ativo) {
        if (norm_(r[0]) === norm_(eu.nome)) throw new Error('Você não pode remover a si mesmo');
        if (r[4] === 'Gestor' && gestores() <= 1) throw new Error('A equipe precisa de pelo menos um Gestor');
        const para = String(p.repassarPara || '');
        const c = cfg_();
        let n = 0;
        lerTarefas_().filter(t => norm_(t.responsavel) === norm_(r[0]) && t.status !== c.COLUNA_FINAL && t.arquivada !== 'SIM')
          .forEach(t => { alterar_(t.id, { responsavel: para }, eu.nome, true); n++; });
        const cs = ss_().getSheetByName(ABA.CATEGORIAS);
        if (cs && cs.getLastRow() > 1) {
          const v = cs.getRange(2, 2, cs.getLastRow() - 1, 1).getDisplayValues();
          let mudou = false;
          v.forEach(x => { if (norm_(x[0]) === norm_(r[0])) { x[0] = para; mudou = true; } });
          if (mudou) cs.getRange(2, 2, v.length, 1).setValues(v);
        }
        sh.getRange(i + 2, 7).setValue('NAO');
        info = `${r[0]} foi removido(a) da equipe.` + (n ? ` ${n} demanda(s) passaram para ${para || 'sem responsável'}.` : '');
      } else {
        sh.getRange(i + 2, 7).setValue('SIM');
        info = `${r[0]} foi reativado(a).`;
      }
    }
    const abertas = {};
    const c = cfg_();
    lerTarefas_().forEach(t => { if (t.responsavel && t.status !== c.COLUNA_FINAL && t.arquivada !== 'SIM') abertas[norm_(t.responsavel)] = (abertas[norm_(t.responsavel)] || 0) + 1; });
    const lista = linhas().filter(r => r[0]).map(r => ({
      nome: r[0], email: r[1], funcao: r[2], capacidade: r[3], perfil: r[4] || 'Membro', pin: r[5], ativo: ativo(r), abertas: abertas[norm_(r[0])] || 0,
      interinoAte: r[4] !== 'Gestor' && /^\d{4}-\d{2}-\d{2}$/.test(String(r[7] || '')) && String(r[7]) >= hoje_() ? String(r[7]) : ''
    }));
    return { lista: lista, info: info };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- limpeza do quadro (Gestor) ---------- */

function limparQuadro_(p, eu) {
  if (String(p.confirmacao || '').trim().toUpperCase() !== 'LIMPAR') throw new Error('Digite LIMPAR para confirmar');
  const escopo = ['tudo', 'trello', 'arquivadas'].indexOf(p.escopo) >= 0 ? p.escopo : 'tudo';
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = ss_();
    const todas = lerTarefas_();
    const alvo = todas.filter(t => escopo === 'trello' ? String(t.origem).indexOf('trello:') === 0
      : escopo === 'arquivadas' ? t.arquivada === 'SIM' : true);
    if (!alvo.length) return { removidas: 0, eventos: 0, backup: '' };

    const carimbo = Utilities.formatDate(new Date(), tz_(), 'dd-MM HH:mm');
    [ABA.TAREFAS, ABA.COMENTARIOS, ABA.HISTORICO].forEach(n => {
      const sh = ss.getSheetByName(n);
      if (sh) sh.copyTo(ss).setName(`Backup ${n} ${carimbo}`).hideSheet();
    });

    let eventos = 0;
    const cal = agenda_();
    if (cal) alvo.forEach(t => {
      if (!t.eventoId) return;
      try { const ev = cal.getEventById(t.eventoId); if (ev) { ev.deleteEvent(); eventos++; } } catch (e) { /* evento já removido */ }
    });

    const ids = {};
    alvo.forEach(t => ids[t.id] = true);
    const manter = todas.filter(t => !ids[t.id]);
    const sh = shTarefas_();
    if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).clearContent();
    if (manter.length) sh.getRange(2, 1, manter.length, CAMPOS.length).setNumberFormat('@').setValues(manter.map(t => linhaDe_(t)[0]));

    [ABA.COMENTARIOS, ABA.HISTORICO].forEach(n => {
      const s = ss.getSheetByName(n);
      if (!s || s.getLastRow() < 2) return;
      const largura = s.getLastColumn();
      const linhas = s.getRange(2, 1, s.getLastRow() - 1, largura).getDisplayValues().filter(r => !ids[r[1]]);
      s.getRange(2, 1, s.getLastRow() - 1, largura).clearContent();
      if (linhas.length) s.getRange(2, 1, linhas.length, largura).setNumberFormat('@').setValues(linhas);
    });

    registrar_('—', eu.nome, 'Limpeza', `${alvo.length} demanda(s) apagada(s) (${escopo}); backup ${carimbo}`);
    return { removidas: alvo.length, eventos: eventos, backup: carimbo };
  } finally {
    lock.releaseLock();
  }
}

function gerarPin_() { return String(Math.floor(1000 + Math.random() * 9000)); }

function enviarAcessoPara_(m) {
  const c = cfg_();
  const url = ScriptApp.getService().getUrl();
  enviarEmail_(m.email, `🔑 Seu acesso ao ${c.NOME_QUADRO}`,
    `<p>Olá, ${esc_(m.nome)}!</p><p>Este é o quadro de demandas da equipe. Lá você acompanha suas entregas, comenta e atualiza o status.</p>
     <p><b>Seu nome:</b> ${esc_(m.nome)}<br><b>Seu PIN:</b> <span style="font-size:20px;letter-spacing:3px">${esc_(m.pin)}</span><br><b>Perfil:</b> ${esc_(m.perfil)}</p>
     <p><a href="${url}" style="background:#1F3A5F;color:#fff;padding:10px 16px;border-radius:6px;text-decoration:none">Abrir o quadro</a></p>
     <p style="color:#6b7280;font-size:13px">No celular, abra o link e use "Adicionar à tela inicial". Não compartilhe seu PIN.</p>`);
}

/* ---------- eventos do Google Agenda exibidos no Calendário do quadro ---------- */

function agendaEventos_(ini, fim) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(ini)) || !/^\d{4}-\d{2}-\d{2}$/.test(String(fim))) return [];
  const c = cfg_();
  const ids = lista_(c.AGENDA_EXIBIR || 'principal', ',');
  const [a1, m1, d1] = ini.split('-').map(Number);
  const [a2, m2, d2] = fim.split('-').map(Number);
  const out = [];
  ids.forEach(id => {
    let cal = null;
    try { cal = id === 'principal' ? CalendarApp.getDefaultCalendar() : CalendarApp.getCalendarById(id); } catch (e) { cal = null; }
    if (!cal) return;
    cal.getEvents(new Date(a1, m1 - 1, d1), new Date(a2, m2 - 1, d2 + 1)).forEach(ev => {
      const titulo = ev.getTitle() || '';
      if (/^(📌|✅) \[/.test(titulo)) return; // entregas do próprio quadro
      const s = ev.getStartTime();
      out.push({
        titulo: titulo, data: Utilities.formatDate(s, tz_(), 'yyyy-MM-dd'),
        hora: ev.isAllDayEvent() ? '' : Utilities.formatDate(s, tz_(), 'HH:mm'),
        local: ev.getLocation() || '', agenda: cal.getName()
      });
    });
  });
  return out.sort((x, y) => (x.data + x.hora) < (y.data + y.hora) ? -1 : 1);
}

function board_(p) {
  const c = cfg_();
  const coment = contarComentarios_();
  const tarefas = lerTarefas_()
    .filter(t => p.arquivadas || t.arquivada !== 'SIM')
    .map(t => { const o = toClient_(t); o.comentarios = coment[t.id] || 0; return o; });
  return {
    config: {
      nome: c.NOME_QUADRO, cor: c.COR_PRINCIPAL || '#1F3A5F',
      colunas: lista_(c.COLUNAS), colunaFinal: c.COLUNA_FINAL,
      prioridades: lista_(c.PRIORIDADES), agenda: c.AGENDA_INTEGRAR === 'SIM',
      urlApp: ScriptApp.getService().getUrl() || ''
    },
    categorias: categorias_(),
    equipe: equipe_().map(m => ({ nome: m.nome, funcao: m.funcao, capacidade: m.capacidade, perfil: m.perfil })),
    tarefas: tarefas,
    agora: agora_()
  };
}

function detalhes_(id) {
  const com = lerAba_(ABA.COMENTARIOS).filter(r => r[1] === id)
    .map(r => ({ id: r[0], autor: r[2], texto: r[3], data: r[4] }));
  const hist = lerAba_(ABA.HISTORICO).filter(r => r[1] === id)
    .map(r => ({ data: r[0], autor: r[2], acao: r[3], detalhe: r[4] })).reverse();
  return { comentarios: com, historico: hist };
}

/* ============================================================
 *  ENTRADA PELO GOOGLE FORMS
 * ============================================================ */

function aoEnviarFormulario(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const origem = e.range ? e.range.getSheet().getName() + '#' + e.range.getRow() : '';
    if (origem && origem.indexOf(ABA.TAREFAS) === 0) return;
    if (origem && jaImportado_(origem, lerTarefas_())) return;
    criarDoFormulario_(e.namedValues, 'Formulário', origem);
  } finally {
    lock.releaseLock();
  }
}

function importarRespostasExistentes() {
  soDono_();
  const aba = abaRespostas_();
  if (!aba) return alerta_('Não encontrei a aba de respostas do formulário.');
  const valores = aba.getDataRange().getDisplayValues();
  const cab = valores[0];
  const tarefasAtuais = lerTarefas_();
  let n = 0;
  for (let i = 1; i < valores.length; i++) {
    const origem = aba.getName() + '#' + (i + 1);
    if (jaImportado_(origem, tarefasAtuais) || !valores[i].join('')) continue;
    const nv = {};
    cab.forEach((h, j) => nv[h] = [valores[i][j]]);
    n += criarDoFormulario_(nv, 'Importação', origem);
  }
  alerta_(`✅ ${n} card(s) criado(s) a partir das respostas.`);
}

function jaImportado_(origem, tarefas) {
  return tarefas.some(t => t.origem === origem || String(t.origem).indexOf(origem + '|') === 0);
}

/** Cria 1 card por categoria envolvida (aba Desdobramento) ou 1 card simples. Retorna quantos criou. */
function criarDoFormulario_(namedValues, autor, origem) {
  const base = dadosDoFormulario_(namedValues);
  const partes = desdobrar_(namedValues);
  if (!partes.length) { criarTarefa_(base, autor, origem); return 1; }
  const hoje = hoje_();
  partes.forEach(pt => {
    const d = Object.assign({}, base);
    d.categoria = pt.categoria;
    d.titulo = `${pt.categoria}: ${base.titulo || 'Nova demanda'}`;
    d.checklistExtra = pt.itens.map(x => '📌 ' + x);
    if (pt.prazo) {
      d.prazo = pt.prazo;
      if (pt.prazo < hoje) { d.prazo = hoje; d.prioridade = 'Urgente'; }
      else if (pt.prazo <= somarDias_(hoje, 1)) d.prioridade = 'Alta';
    } else d.prazo = '';
    d.etiquetas = partes.length > 1 ? `evento (${partes.length} áreas)` : '';
    criarTarefa_(d, autor, origem ? origem + '|' + pt.categoria : '');
  });
  return partes.length;
}

function desdobrar_(namedValues) {
  const regras = lerAba_(ABA.DESDOBRAMENTO).filter(r => r[0] === 'SIM' && r[1] && r[3]);
  if (!regras.length) return [];
  const perguntas = Object.keys(namedValues);
  const resp = chave => {
    const k = perguntas.find(q => norm_(q).indexOf(norm_(chave)) >= 0);
    return k ? (namedValues[k] || []).join(', ').trim() : '';
  };
  const porCat = {};
  regras.forEach(([, pergunta, contem, categoria, prazoPergunta, diasAntes]) => {
    const r = resp(pergunta);
    if (!r) return;
    const opcoes = dividirOpcoes_(r);
    const casadas = contem === '*' ? opcoes : opcoes.filter(o => norm_(o).indexOf(norm_(contem)) >= 0);
    if (!casadas.length) return;
    const p = porCat[categoria] = porCat[categoria] || { categoria: categoria, itens: [], prazo: '' };
    casadas.forEach(o => { if (p.itens.indexOf(o) < 0) p.itens.push(o); });
    const base = prazoPergunta ? normalizarData_(resp(prazoPergunta)) : '';
    if (base) {
      const pz = somarDias_(base, -Number(diasAntes || 0));
      if (!p.prazo || pz < p.prazo) p.prazo = pz;
    }
  });
  // remove itens genéricos ("Arte") quando já existe um mais específico ("Arte para Redes Sociais")
  Object.values(porCat).forEach(p => {
    p.itens = p.itens.filter(x => !p.itens.some(y => y !== x && norm_(y).indexOf(norm_(x)) === 0));
  });
  // sem data de divulgação, usa a data do evento como teto
  const evento = normalizarData_(resp('Data de Início do Evento'));
  Object.values(porCat).forEach(p => { if (!p.prazo && evento) p.prazo = somarDias_(evento, -1); });
  return Object.values(porCat);
}

/** Divide "A, B (x, y), C" em ["A", "B (x, y)", "C"], respeitando parênteses. */
function dividirOpcoes_(r) {
  const out = []; let atual = '', nivel = 0;
  for (const ch of String(r)) {
    if (ch === '(') nivel++;
    if (ch === ')') nivel = Math.max(0, nivel - 1);
    if (ch === ',' && nivel === 0) { if (atual.trim()) out.push(atual.trim()); atual = ''; } else atual += ch;
  }
  if (atual.trim()) out.push(atual.trim());
  return out;
}

function dadosDoFormulario_(namedValues) {
  const mapa = mapeamento_();
  const ignorar = /^(carimbo de data\/hora|timestamp|marca temporal)$/i;
  const briefing = [], anexos = [];
  const dados = {};
  Object.keys(namedValues).forEach(pergunta => {
    if (ignorar.test(pergunta.trim())) return;
    const resposta = (namedValues[pergunta] || []).join(', ').trim();
    if (!resposta) return;
    const urls = resposta.match(/https:\/\/drive\.google\.com\/[^\s,]+/g);
    if (urls) urls.forEach((u, i) => anexos.push({ nome: pergunta + (urls.length > 1 ? ' ' + (i + 1) : ''), url: u }));
    else briefing.push({ p: pergunta, r: resposta });
    const campo = mapa[pergunta.trim().toLowerCase()] || adivinharCampo_(pergunta, mapa);
    if (campo && !dados[campo]) dados[campo] = resposta;
  });
  dados.briefing = briefing;
  dados.anexos = anexos;
  if (dados.prazo) dados.prazo = normalizarData_(dados.prazo);
  if (dados.solicitanteComplemento) dados.solicitante = `${dados.solicitante || ''} (${dados.solicitanteComplemento})`.trim();
  return dados;
}

function adivinharCampo_(pergunta, mapa) {
  const usados = new Set(Object.values(mapa));
  const q = semAcento_(pergunta.toLowerCase());
  const regras = [
    ['emailSolicitante', /e-?mail/],
    ['contato', /whats|telefone|celular|contato/],
    ['prazo', /prazo|data (de|da) entrega|para quando|data do evento|data limite/],
    ['categoria', /tipo de demanda|categoria|tipo de servico|o que voce precisa/],
    ['titulo', /titulo|nome do evento|nome da demanda|assunto/],
    ['solicitante', /seu nome|solicitante|ministerio|rede|departamento|nome completo/],
    ['prioridade', /prioridade|urgencia/]
  ];
  for (const [campo, re] of regras) if (!usados.has(campo) && re.test(q)) return campo;
  return '';
}

/* ============================================================
 *  IMPORTAÇÃO DO TRELLO
 * ============================================================ */

function abrirImportacaoTrello() {
  const url = ScriptApp.getService().getUrl();
  alerta_(url ? 'Abra o quadro como Gestor e use o botão "Importar do Trello" (ou acesse ' + url + '?pagina=importar).' : 'Implante o app da web primeiro.');
}

/** Dados para a tela de mapeamento. */
function trelloContexto_() {
  const c = cfg_();
  return {
    colunas: lista_(c.COLUNAS), colunaFinal: c.COLUNA_FINAL,
    equipe: equipe_().map(m => m.nome), categorias: categorias_().map(x => x.nome),
    jaImportados: lerTarefas_().filter(t => String(t.origem).indexOf('trello:') === 0).length,
    urlApp: ScriptApp.getService().getUrl() || ''
  };
}

/** Busca o quadro completo pela API do Trello (sem limite de comentários). */
function trelloBuscarApi_(key, token, linkQuadro) {
  const m = String(linkQuadro || '').match(/trello\.com\/b\/([A-Za-z0-9]+)/);
  const idQuadro = m ? m[1] : String(linkQuadro || '').trim();
  if (!key || !token || !idQuadro) throw new Error('Preencha a chave, o token e o link do quadro.');
  const base = 'https://api.trello.com/1';
  const auth = `key=${encodeURIComponent(key)}&token=${encodeURIComponent(token)}`;
  const get = url => {
    const r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (r.getResponseCode() !== 200) throw new Error(`Trello respondeu ${r.getResponseCode()}: ${r.getContentText().slice(0, 150)}`);
    return JSON.parse(r.getContentText());
  };
  const board = get(`${base}/boards/${idQuadro}?lists=all&cards=all&card_attachments=true&checklists=all&members=all&labels=all&fields=name,url&${auth}`);
  const acoes = [];
  let antes = '';
  for (let i = 0; i < 50; i++) {
    const lote = get(`${base}/boards/${idQuadro}/actions?filter=commentCard&limit=1000${antes ? '&before=' + antes : ''}&${auth}`);
    acoes.push.apply(acoes, lote);
    if (lote.length < 1000) break;
    antes = lote[lote.length - 1].id;
  }
  board.actions = acoes;
  return board;
}

/**
 * Importa os cards já normalizados pelo diálogo.
 * d = { cards:[...], comentarios:[...], mapaListas:{idLista:status|''}, mapaMembros:{idMembro:nome},
 *       listas:{idLista:nome}, incluirArquivados:bool, criarEventos:bool }
 */
function trelloImportar_(d) {
  const inicio = Date.now();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const c = cfg_();
    const colunas = lista_(c.COLUNAS);
    const cats = categorias_();
    const ss = ss_();
    const existentes = {};
    lerTarefas_().forEach(t => { if (String(t.origem).indexOf('trello:') === 0) existentes[t.origem] = t.id; });

    const tarefas = [];
    const idPorCard = {};
    let puladas = 0, ignoradas = 0, eventos = 0, eventosPendentes = 0;

    (d.cards || []).forEach((card, i) => {
      const origem = 'trello:' + card.id;
      if (existentes[origem]) { puladas++; return; }
      const status = d.mapaListas[card.idList];
      if (!status || colunas.indexOf(status) < 0) { ignoradas++; return; }
      if (card.closed && !d.incluirArquivados) { ignoradas++; return; }

      // categoria: etiqueta que bate com uma categoria; senão, o nome da lista
      let categoria = '';
      const etiquetas = [];
      (card.labels || []).forEach(l => {
        const m = !categoria && combinarCategoria_(l, cats);
        if (m) categoria = m.nome; else if (l) etiquetas.push(l);
      });
      if (!categoria) { const m = combinarCategoria_(d.listas[card.idList], cats); if (m) categoria = m.nome; }

      const responsaveis = (card.idMembers || []).map(id => d.mapaMembros[id]).filter(String);
      const criadoEm = Utilities.formatDate(new Date(parseInt(card.id.substring(0, 8), 16) * 1000), tz_(), 'yyyy-MM-dd HH:mm');
      const final = status === c.COLUNA_FINAL;
      const anexos = (card.anexos || []).slice();
      if (card.url) anexos.push({ nome: 'Card original no Trello', url: card.url });

      const t = {
        id: 'T' + Utilities.formatDate(new Date(), tz_(), 'yyMMddHHmm') + String(1000 + i).slice(-4) + Math.floor(Math.random() * 9),
        criadoEm: criadoEm,
        titulo: card.name || '(sem título)',
        categoria: categoria,
        status: status,
        responsavel: responsaveis[0] || '',
        inicio: card.start ? isoLocal_(card.start) : criadoEm.slice(0, 10),
        prazo: card.due ? isoLocal_(card.due) : '',
        prioridade: c.PRIORIDADE_PADRAO || 'Normal',
        etiquetas: etiquetas.concat(responsaveis.slice(1).map(n => '+' + n)).join(', '),
        solicitante: '', contato: '', emailSolicitante: '',
        briefing: briefingDaDescricao_(card.desc),
        anexos: anexos,
        checklist: (card.checklist || []).map(x => ({ id: uid_(), texto: x.texto, feito: !!x.feito })),
        estimativa: '', tempoMin: '0', timerInicio: '', timerAutor: '', eventoId: '',
        arquivada: card.closed ? 'SIM' : 'NAO',
        concluidaEm: final ? (card.dateLastActivity ? Utilities.formatDate(new Date(card.dateLastActivity), tz_(), 'yyyy-MM-dd HH:mm') : agora_()) : '',
        atualizadoEm: agora_(),
        ordem: String(Number(card.pos) || (i + 1) * 1000),
        origem: origem,
        _lista: d.listas[card.idList] || ''
      };
      // campos "Solicitante/WhatsApp/E-mail" que vieram do Forms via Make
      const campo = rx => (t.briefing.find(b => rx.test(norm_(b.p))) || {}).r || '';
      t.solicitante = campo(/solicitante|seu nome|nome completo|ministerio|rede/);
      t.contato = campo(/whats|telefone|celular|contato/);
      t.emailSolicitante = (campo(/e-?mail/).match(/[^\s@]+@[^\s@]+\.[^\s@]+/) || [''])[0];

      if (d.criarEventos && t.prazo && !final && t.arquivada !== 'SIM') {
        if (Date.now() - inicio < 270000) { sincronizarAgenda_(t); if (t.eventoId) eventos++; }
        else eventosPendentes++;
      }
      idPorCard[card.id] = t.id;
      tarefas.push(t);
    });

    if (tarefas.length) {
      const sh = shTarefas_();
      const linha = sh.getLastRow() + 1;
      sh.getRange(linha, 1, tarefas.length, CAMPOS.length).setNumberFormat('@')
        .setValues(tarefas.map(t => linhaDe_(t)[0]));
    }

    const comentarios = (d.comentarios || []).filter(x => idPorCard[x.idCard]).map(x => [
      uid_(), idPorCard[x.idCard], x.autor || 'Trello', x.texto || '',
      Utilities.formatDate(new Date(x.data), tz_(), 'yyyy-MM-dd HH:mm')
    ]).sort((a, b) => a[4] < b[4] ? -1 : 1);
    if (comentarios.length) {
      const shC = ss.getSheetByName(ABA.COMENTARIOS);
      shC.getRange(shC.getLastRow() + 1, 1, comentarios.length, 5).setNumberFormat('@').setValues(comentarios);
    }

    if (tarefas.length) {
      const shH = ss.getSheetByName(ABA.HISTORICO);
      const agora = agora_();
      shH.getRange(shH.getLastRow() + 1, 1, tarefas.length, 5).setNumberFormat('@')
        .setValues(tarefas.map(t => [agora, t.id, 'Importação Trello', 'Importou', `lista "${t._lista}" → ${t.status}`]));
    }

    return { importadas: tarefas.length, puladas: puladas, ignoradas: ignoradas, comentarios: comentarios.length, eventos: eventos, eventosPendentes: eventosPendentes };
  } finally {
    lock.releaseLock();
  }
}

/** Transforma "Campo: valor" (padrão do Make) em itens de briefing. */
function briefingDaDescricao_(desc) {
  desc = String(desc || '').trim();
  if (!desc) return [];
  const itens = [], soltas = [];
  desc.split(/\r?\n/).forEach(linha => {
    const m = linha.match(/^\s*(?:[-*•]\s*)?[*_]{0,2}([^:*_\n]{2,60}?)[*_]{0,2}\s*:\s*[*_]{0,2}\s*(.+?)\s*[*_]{0,2}$/);
    if (m && !/^https?$/i.test(m[1].trim())) itens.push({ p: m[1].trim(), r: m[2].trim() });
    else if (linha.trim()) soltas.push(linha.trim());
  });
  if (itens.length < 2) return [{ p: 'Descrição (Trello)', r: desc }];
  if (soltas.length) itens.push({ p: 'Observações (Trello)', r: soltas.join('\n') });
  return itens;
}

function isoLocal_(dataIso) {
  const d = new Date(dataIso);
  return isNaN(d) ? '' : Utilities.formatDate(d, tz_(), 'yyyy-MM-dd');
}

/* ============================================================
 *  REGRAS DE NEGÓCIO
 * ============================================================ */

function criarTarefa_(d, autor, origem) {
  const c = cfg_();
  const cats = categorias_();
  const cat = combinarCategoria_(d.categoria, cats);
  const colunas = lista_(c.COLUNAS);
  const hoje = hoje_();
  const t = {
    id: 'T' + Utilities.formatDate(new Date(), tz_(), 'yyMMddHHmmss') + Math.floor(Math.random() * 90 + 10),
    criadoEm: agora_(),
    titulo: d.titulo || (cat ? cat.nome + ' — ' : '') + (d.solicitante || 'Nova demanda'),
    categoria: cat ? cat.nome : (d.categoria || ''),
    status: d.status && colunas.includes(d.status) ? d.status : colunas[0],
    responsavel: d.responsavel || (cat ? cat.responsavel : ''),
    inicio: d.inicio || hoje,
    prazo: d.prazo || (cat && cat.sla ? somarDias_(hoje, cat.sla) : ''),
    prioridade: d.prioridade || c.PRIORIDADE_PADRAO || 'Normal',
    etiquetas: d.etiquetas || '',
    solicitante: d.solicitante || '',
    contato: d.contato || '',
    emailSolicitante: d.emailSolicitante || '',
    briefing: d.briefing || (d.descricao ? [{ p: 'Descrição', r: d.descricao }] : []),
    anexos: d.anexos || [],
    checklist: (d.checklistExtra || []).concat(cat ? cat.checklist : []).map(txt => ({ id: uid_(), texto: txt, feito: false })),
    estimativa: d.estimativa || '',
    tempoMin: '0', timerInicio: '', timerAutor: '', eventoId: '',
    arquivada: 'NAO', concluidaEm: '', atualizadoEm: agora_(),
    ordem: String(Date.now()), origem: origem || ''
  };
  const efeitos = executarAutomacoes_('Tarefa criada', t, {});
  sincronizarAgenda_(t);
  inserirTarefa_(t);
  registrar_(t.id, autor, 'Criou', `via ${origem && origem.indexOf('#') > 0 ? 'formulário' : origem || autor}`);
  notificarNova_(t);
  efeitos.forEach(f => f());
  return t;
}

/** Aplica alterações, registra histórico, roda automações, sincroniza agenda e notifica. */
function alterar_(id, campos, autor, semAutomacao) {
  const t = buscarTarefa_(id);
  if (!t) throw new Error('Tarefa não encontrada');
  const antes = Object.assign({}, t);
  const c = cfg_();
  autor = autor || 'Alguém';
  const rotulos = { titulo: 'Título', categoria: 'Categoria', status: 'Status', responsavel: 'Responsável', prazo: 'Prazo', inicio: 'Início', prioridade: 'Prioridade', etiquetas: 'Etiquetas', estimativa: 'Estimativa', arquivada: 'Arquivada' };
  Object.keys(campos).forEach(k => {
    if (!CAMPOS.some(cc => cc[0] === k)) return;
    let v = campos[k];
    if (v === undefined || v === null) v = '';
    if (CAMPOS_JSON.includes(k)) {
      if (JSON.stringify(t[k]) === JSON.stringify(v)) return;
      t[k] = v;
      if (k === 'checklist') {
        const feitos = v.filter(i => i.feito).length;
        registrar_(id, autor, 'Checklist', `${feitos}/${v.length} concluídos`);
      } else registrar_(id, autor, 'Anexos', `${v.length} anexo(s)`);
      return;
    }
    v = String(v);
    if (t[k] === v) return;
    t[k] = v;
    if (rotulos[k]) registrar_(id, autor, rotulos[k], `${antes[k] || '—'} → ${v || '—'}`);
  });

  const final = c.COLUNA_FINAL;
  if (t.status !== antes.status) {
    t.concluidaEm = t.status === final ? agora_() : '';
    if (t.status === final && t.timerInicio) pararTimerInterno_(t);
  }
  t.atualizadoEm = agora_();

  let efeitos = [];
  if (!semAutomacao) {
    if (t.status !== antes.status) efeitos = efeitos.concat(executarAutomacoes_('Status mudou para', t, antes));
    if (t.responsavel !== antes.responsavel) efeitos = efeitos.concat(executarAutomacoes_('Responsável mudou', t, antes));
  }

  const agendaMudou = ['titulo', 'prazo', 'responsavel', 'status', 'arquivada', 'categoria'].some(k => t[k] !== antes[k]);
  if (agendaMudou) sincronizarAgenda_(t);
  salvarTarefa_(t);

  if (t.responsavel !== antes.responsavel && t.responsavel) notificarResponsavel_(t, 'Você recebeu uma demanda');
  if (t.status !== antes.status) notificarStatus_(t, antes.status);
  efeitos.forEach(f => f());
  return t;
}

function comentar_(id, autor, texto) {
  texto = String(texto || '').trim();
  if (!texto) throw new Error('Comentário vazio');
  const t = buscarTarefa_(id);
  const data = agora_();
  const cid = uid_();
  ss_().getSheetByName(ABA.COMENTARIOS).appendRow([cid, id, autor || 'Alguém', texto, data]);
  registrar_(id, autor, 'Comentou', texto.slice(0, 80));
  // @menções
  const equipe = equipe_();
  const mencionados = equipe.filter(m => m.email && new RegExp('@' + escRe_(m.nome.split(' ')[0]), 'i').test(texto));
  mencionados.forEach(m => enviarEmail_(m.email, `💬 ${autor} mencionou você em "${t.titulo}"`,
    `<p><b>${esc_(autor)}</b> escreveu:</p><blockquote>${esc_(texto)}</blockquote>${blocoTarefa_(t)}`));
  if (t.responsavel && t.responsavel !== autor && !mencionados.some(m => m.nome === t.responsavel)) {
    const r = membro_(t.responsavel);
    if (r && r.email) enviarEmail_(r.email, `💬 Novo comentário em "${t.titulo}"`,
      `<p><b>${esc_(autor)}</b> comentou:</p><blockquote>${esc_(texto)}</blockquote>${blocoTarefa_(t)}`);
  }
  return { id: cid, autor: autor, texto: texto, data: data };
}

function timer_(id, modo, autor) {
  const t = buscarTarefa_(id);
  if (modo === 'iniciar') {
    if (t.timerInicio) return t;
    t.timerInicio = String(Date.now());
    t.timerAutor = autor || '';
    registrar_(id, autor, 'Timer', 'iniciado');
  } else {
    pararTimerInterno_(t);
    registrar_(id, autor, 'Timer', 'parado');
  }
  t.atualizadoEm = agora_();
  salvarTarefa_(t);
  return t;
}

function pararTimerInterno_(t) {
  if (!t.timerInicio) return;
  const min = Math.max(1, Math.round((Date.now() - Number(t.timerInicio)) / 60000));
  t.tempoMin = String(Number(t.tempoMin || 0) + min);
  t.timerInicio = '';
  t.timerAutor = '';
}

function tempoManual_(id, minutos, autor) {
  const t = buscarTarefa_(id);
  const m = Math.round(Number(minutos) || 0);
  t.tempoMin = String(Math.max(0, Number(t.tempoMin || 0) + m));
  registrar_(id, autor, 'Tempo', `${m > 0 ? '+' : ''}${m} min lançados`);
  salvarTarefa_(t);
  return t;
}

/* ============================================================
 *  AUTOMAÇÕES (estilo Butler / ClickUp Automations)
 * ============================================================ */

function executarAutomacoes_(gatilho, t, antes) {
  const efeitos = [];
  const regras = lerAba_(ABA.AUTOMACOES).filter(r => r[0] === 'SIM' && r[1] === gatilho);
  regras.forEach(r => {
    const [, , valorGatilho, seCampo, igualA, acao, valorAcao] = r;
    if (gatilho === 'Status mudou para' && valorGatilho && norm_(valorGatilho) !== norm_(t.status)) return;
    if (gatilho === 'Responsável mudou' && valorGatilho && norm_(valorGatilho) !== norm_(t.responsavel)) return;
    if (seCampo && !condicao_(t, seCampo, igualA)) return;
    const rotulo = `${gatilho}${valorGatilho ? ' ' + valorGatilho : ''} → ${acao} ${valorAcao}`;
    switch (acao) {
      case 'Mover para': t.status = valorAcao; break;
      case 'Atribuir a': t.responsavel = valorAcao; break;
      case 'Definir prioridade': t.prioridade = valorAcao; break;
      case 'Adicionar etiqueta': {
        const tags = lista_(t.etiquetas, ',');
        if (!tags.includes(valorAcao)) tags.push(valorAcao);
        t.etiquetas = tags.join(', ');
        break;
      }
      case 'Adicionar item de checklist':
        t.checklist = (t.checklist || []).concat([{ id: uid_(), texto: valorAcao, feito: false }]);
        break;
      case 'Notificar e-mail': {
        const destino = valorAcao === 'coordenacao' ? emailsCoordenacao_()
          : valorAcao === 'responsavel' ? (membro_(t.responsavel) || {}).email
          : (membro_(valorAcao) || {}).email || valorAcao;
        if (destino) efeitos.push(() => enviarEmail_(destino, `⚡ ${t.titulo} — ${gatilho.toLowerCase()} ${valorGatilho || ''}`.trim(),
          `<p>Automação disparada: <b>${esc_(rotulo)}</b></p>${blocoTarefa_(t)}`));
        break;
      }
      case 'Enviar webhook':
        efeitos.push(() => webhook_('automacao', t, { regra: rotulo, url: valorAcao }));
        break;
    }
    if (t.id) efeitos.push(() => registrar_(t.id, '⚡ Automação', 'Regra', rotulo));
  });
  return efeitos;
}

function condicao_(t, campo, valor) {
  const v = norm_(valor);
  switch (norm_(campo)) {
    case 'categoria': return norm_(t.categoria) === v;
    case 'prioridade': return norm_(t.prioridade) === v;
    case 'responsavel': return norm_(t.responsavel) === v;
    case 'status': return norm_(t.status) === v;
    case 'etiqueta': return lista_(t.etiquetas, ',').map(norm_).includes(v);
    default: return true;
  }
}

/* ============================================================
 *  ROTINA DIÁRIA: atrasos, recorrentes, arquivamento
 * ============================================================ */

function rotinaDiaria() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    gerarRecorrentes_();
    verificarAtrasos_();
    arquivarAntigas_();
  } finally {
    lock.releaseLock();
  }
}

function verificarAtrasos_() {
  const c = cfg_();
  const hoje = hoje_();
  const abertas = lerTarefas_().filter(t => t.arquivada !== 'SIM' && t.status !== c.COLUNA_FINAL);
  const atrasadas = abertas.filter(t => t.prazo && t.prazo < hoje);
  const vencemHoje = abertas.filter(t => t.prazo === hoje);
  atrasadas.forEach(t => {
    const efeitos = executarAutomacoes_('Prazo vencido', t, {});
    if (efeitos.length) { salvarTarefa_(t); efeitos.forEach(f => f()); }
  });
  const porPessoa = {};
  atrasadas.concat(vencemHoje).forEach(t => (porPessoa[t.responsavel || ''] = porPessoa[t.responsavel || ''] || []).push(t));
  Object.keys(porPessoa).forEach(nome => {
    const m = membro_(nome);
    if (!m || !m.email) return;
    enviarEmail_(m.email, `⏰ Você tem ${porPessoa[nome].length} entrega(s) vencendo ou atrasada(s)`, listaHtml_(porPessoa[nome], hoje));
  });
  const coord = emailsCoordenacao_();
  if (coord && (atrasadas.length || vencemHoje.length)) {
    enviarEmail_(coord, `📊 Resumo do dia — ${atrasadas.length} atrasada(s), ${vencemHoje.length} vencendo hoje`,
      listaHtml_(atrasadas.concat(vencemHoje), hoje));
  }
}

function gerarRecorrentes_() {
  const sh = ss_().getSheetByName(ABA.RECORRENTES);
  if (!sh) return;
  const dados = sh.getDataRange().getDisplayValues();
  const hoje = hoje_();
  const d = new Date();
  const diaSemana = ['dom', 'seg', 'ter', 'qua', 'qui', 'sex', 'sab'][d.getDay()];
  for (let i = 1; i < dados.length; i++) {
    const [ativa, titulo, categoria, responsavel, freq, dia, prazoDias, descricao, ultima] = dados[i];
    if (ativa !== 'SIM' || !titulo || ultima === hoje) continue;
    let gerar = false;
    if (freq === 'Diária') gerar = true;
    if (freq === 'Semanal') gerar = semAcento_(String(dia).toLowerCase()).slice(0, 3) === diaSemana;
    if (freq === 'Mensal') gerar = Number(dia) === d.getDate();
    if (!gerar) continue;
    criarTarefa_({
      titulo: titulo + ' — ' + Utilities.formatDate(d, tz_(), 'dd/MM'),
      categoria: categoria, responsavel: responsavel,
      prazo: prazoDias ? somarDias_(hoje, Number(prazoDias)) : '',
      descricao: descricao, solicitante: 'Recorrente'
    }, '🔁 Recorrência', 'recorrente');
    sh.getRange(i + 1, 9).setNumberFormat('@').setValue(hoje);
  }
}

function arquivarAntigas_() {
  const c = cfg_();
  const dias = Number(c.ARQUIVAR_APOS_DIAS || 0);
  if (!dias) return;
  const limite = somarDias_(hoje_(), -dias);
  lerTarefas_().filter(t => t.arquivada !== 'SIM' && t.status === c.COLUNA_FINAL && t.concluidaEm && t.concluidaEm.slice(0, 10) <= limite)
    .forEach(t => { t.arquivada = 'SIM'; salvarTarefa_(t); registrar_(t.id, '🗄️ Sistema', 'Arquivada', `concluída há mais de ${dias} dias`); });
}

/* ============================================================
 *  GOOGLE AGENDA
 * ============================================================ */

function agenda_() {
  const c = cfg_();
  if (c.AGENDA_INTEGRAR !== 'SIM') return null;
  return c.AGENDA_ID ? CalendarApp.getCalendarById(c.AGENDA_ID) : CalendarApp.getDefaultCalendar();
}

function sincronizarAgenda_(t) {
  try {
    const cal = agenda_();
    if (!cal) return;
    const c = cfg_();
    let ev = null;
    if (t.eventoId) { try { ev = cal.getEventById(t.eventoId); } catch (err) { ev = null; } }
    if (!t.prazo || t.arquivada === 'SIM') {
      if (ev) ev.deleteEvent();
      t.eventoId = '';
      return;
    }
    const concluida = t.status === c.COLUNA_FINAL;
    const titulo = `${concluida ? '✅' : '📌'} [${t.categoria || 'Demanda'}] ${t.titulo}`;
    const url = ScriptApp.getService().getUrl() || '';
    const descricao = [
      `Responsável: ${t.responsavel || '—'}`,
      `Status: ${t.status}`,
      `Prioridade: ${t.prioridade}`,
      `Solicitante: ${t.solicitante || '—'} ${t.contato ? '(' + t.contato + ')' : ''}`,
      '',
      ...(t.briefing || []).slice(0, 8).map(b => `• ${b.p}: ${b.r}`),
      '',
      url ? `Abrir no quadro: ${url}#${t.id}` : ''
    ].join('\n');
    const [a, m, d] = t.prazo.split('-').map(Number);
    const hora = (c.AGENDA_HORA_ENTREGA || '').match(/^(\d{1,2}):(\d{2})$/);
    const ini = hora ? new Date(a, m - 1, d, Number(hora[1]), Number(hora[2])) : new Date(a, m - 1, d);
    const fim = hora ? new Date(ini.getTime() + 30 * 60000) : null;
    const resp = membro_(t.responsavel);
    const convidado = c.AGENDA_CONVIDAR_RESPONSAVEL === 'SIM' && resp && resp.email ? resp.email.toLowerCase() : '';

    if (!ev) {
      const op = { description: descricao, sendInvites: c.AGENDA_CONVITE_POR_EMAIL === 'SIM' };
      if (convidado) op.guests = convidado;
      ev = hora ? cal.createEvent(titulo, ini, fim, op) : cal.createAllDayEvent(titulo, ini, op);
      ev.addPopupReminder(hora ? 24 * 60 : 15 * 60);
      t.eventoId = ev.getId();
    } else {
      ev.setTitle(titulo);
      ev.setDescription(descricao);
      if (hora) ev.setTime(ini, fim); else ev.setAllDayDate(ini);
      ev.getGuestList().forEach(g => { if (g.getEmail().toLowerCase() !== convidado) ev.removeGuest(g.getEmail()); });
      if (convidado && !ev.getGuestList().some(g => g.getEmail().toLowerCase() === convidado)) ev.addGuest(convidado);
    }
    ev.setColor(concluida ? CalendarApp.EventColor.GREEN
      : t.prioridade === 'Urgente' ? CalendarApp.EventColor.RED
      : t.prioridade === 'Alta' ? CalendarApp.EventColor.ORANGE : CalendarApp.EventColor.BLUE);
  } catch (err) {
    registrar_(t.id, '🗓️ Agenda', 'Erro', String(err.message || err).slice(0, 200));
  }
}

function sincronizarTodasAgenda() {
  soDono_();
  lerTarefas_().forEach(t => { sincronizarAgenda_(t); salvarTarefa_(t); });
  alerta_('✅ Agenda sincronizada.');
}

/* ============================================================
 *  NOTIFICAÇÕES
 * ============================================================ */

/** E-mails da coordenação: o configurado + gestores interinos ativos. */
function emailsCoordenacao_() {
  const lista = lista_(cfg_().EMAIL_COORDENACAO, ',');
  equipe_().forEach(m => { if (m.interinoAte && m.email && lista.indexOf(m.email) < 0) lista.push(m.email); });
  return lista.join(', ');
}

function notificarNova_(t) {
  const c = cfg_();
  if (c.NOTIFICAR_RESPONSAVEL === 'SIM' && t.responsavel) notificarResponsavel_(t, 'Nova demanda para você');
  const coord = emailsCoordenacao_();
  if (coord) enviarEmail_(coord, `🆕 Nova demanda: ${t.titulo}`, blocoTarefa_(t));
  webhook_('nova_tarefa', t);
}

function notificarResponsavel_(t, assunto) {
  const m = membro_(t.responsavel);
  if (!m || !m.email || cfg_().NOTIFICAR_RESPONSAVEL !== 'SIM') return;
  enviarEmail_(m.email, `📥 ${assunto}: ${t.titulo}`, blocoTarefa_(t));
  webhook_('atribuida', t, { telefoneResponsavel: m.telefone || '' });
}

function notificarStatus_(t, anterior) {
  const c = cfg_();
  if (c.NOTIFICAR_SOLICITANTE === 'SIM' && /@/.test(t.emailSolicitante)) {
    enviarEmail_(t.emailSolicitante, `🔄 Sua demanda "${t.titulo}" agora está em: ${t.status}`,
      `<p>Olá${t.solicitante ? ', ' + esc_(t.solicitante) : ''}!</p><p>A demanda <b>${esc_(t.titulo)}</b> passou de <i>${esc_(anterior)}</i> para <b>${esc_(t.status)}</b>.</p><p>Equipe ${esc_(c.NOME_QUADRO)}</p>`);
  }
  webhook_('status', t, { anterior: anterior });
}

function webhook_(evento, t, extra) {
  const url = (extra && extra.url) || cfg_().WEBHOOK_NOTIFICACAO;
  if (!url) return;
  try {
    UrlFetchApp.fetch(url, {
      method: 'post', contentType: 'application/json', muteHttpExceptions: true,
      payload: JSON.stringify({ evento: evento, quadro: cfg_().NOME_QUADRO, link: ScriptApp.getService().getUrl(), tarefa: toClient_(t), extra: extra || {} })
    });
  } catch (err) { /* não bloqueia o fluxo */ }
}

function enviarEmail_(para, assunto, corpoHtml) {
  if (!para) return;
  try {
    const c = cfg_();
    MailApp.sendEmail({
      to: para, subject: assunto, name: c.NOME_QUADRO || 'Quadro de Demandas',
      htmlBody: `<div style="font-family:Arial,sans-serif;max-width:560px;color:#1f2937">${corpoHtml}
        <p style="color:#9ca3af;font-size:12px;margin-top:24px">${esc_(c.NOME_QUADRO)} · Quadro de Demandas</p></div>`
    });
  } catch (err) { console.warn('E-mail não enviado', err); }
}

function blocoTarefa_(t) {
  const url = ScriptApp.getService().getUrl();
  const linhas = [
    ['Categoria', t.categoria], ['Responsável', t.responsavel], ['Prazo', dataBr_(t.prazo)],
    ['Prioridade', t.prioridade], ['Status', t.status], ['Solicitante', [t.solicitante, t.contato].filter(String).join(' · ')]
  ].filter(l => l[1]);
  const brief = (t.briefing || []).slice(0, 10).map(b => `<li><b>${esc_(b.p)}:</b> ${esc_(b.r)}</li>`).join('');
  return `<h2 style="margin:0 0 8px">${esc_(t.titulo)}</h2>
    <table style="border-collapse:collapse">${linhas.map(l => `<tr><td style="padding:2px 12px 2px 0;color:#6b7280">${l[0]}</td><td><b>${esc_(l[1])}</b></td></tr>`).join('')}</table>
    ${brief ? `<ul style="padding-left:18px">${brief}</ul>` : ''}
    ${url ? `<p><a href="${url}#${t.id}" style="background:#1F3A5F;color:#fff;padding:10px 16px;border-radius:6px;text-decoration:none">Abrir no quadro</a></p>` : ''}`;
}

function listaHtml_(lista, hoje) {
  const url = ScriptApp.getService().getUrl();
  return `<ul>${lista.map(t => `<li><b>${esc_(t.titulo)}</b> — ${esc_(t.responsavel || 'sem responsável')} — prazo ${dataBr_(t.prazo)} ${t.prazo < hoje ? '<span style="color:#dc2626">(atrasada)</span>' : '(hoje)'}</li>`).join('')}</ul>
    ${url ? `<p><a href="${url}">Abrir quadro</a></p>` : ''}`;
}

/* ============================================================
 *  ACESSO À PLANILHA
 * ============================================================ */

let _cfgCache = null;
function cfg_() {
  if (_cfgCache) return _cfgCache;
  const o = {};
  CONFIG_PADRAO.forEach(r => o[r[0]] = r[1]);
  const sh = ss_().getSheetByName(ABA.CONFIG);
  if (sh) sh.getDataRange().getDisplayValues().slice(1).forEach(r => { if (r[0]) o[r[0].trim()] = String(r[1]).trim(); });
  return (_cfgCache = o);
}

function lerAba_(nome) {
  const sh = ss_().getSheetByName(nome);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getDisplayValues();
}

function categorias_() {
  return lerAba_(ABA.CATEGORIAS).filter(r => r[0]).map(r => ({
    nome: r[0], responsavel: r[1], cor: r[2] || '#64748B', sla: Number(r[3]) || 0,
    checklist: lista_(r[4], ';')
  }));
}

function equipe_(todos) {
  return lerAba_(ABA.EQUIPE).filter(r => r[0]).map(r => ({
    nome: r[0], email: r[1], funcao: r[2], capacidade: Number(r[3]) || 0,
    perfil: ['Gestor', 'Membro', 'Leitor'].indexOf(r[4]) >= 0 ? r[4] : 'Membro', pin: r[5] || '',
    ativo: String(r[6] || 'SIM').toUpperCase() !== 'NAO',
    interinoAte: r[4] !== 'Gestor' && /^\d{4}-\d{2}-\d{2}$/.test(String(r[7] || '')) && String(r[7]) >= hoje_() ? String(r[7]) : ''
  })).map(m => { m.perfilBase = m.perfil; if (m.interinoAte) m.perfil = 'Gestor'; return m; })
    .filter(m => todos || m.ativo);
}

function membro_(nome) {
  if (!nome) return null;
  return equipe_().find(m => norm_(m.nome) === norm_(nome)) || null;
}

function mapeamento_() {
  const o = {};
  lerAba_(ABA.MAPEAMENTO).forEach(r => { if (r[0] && r[1]) o[r[1].trim().toLowerCase()] = r[0].trim(); });
  return o;
}

function combinarCategoria_(valor, cats) {
  if (!valor) return null;
  const v = norm_(valor);
  return cats.find(c => norm_(c.nome) === v) || (v.length < 4 ? null : cats.find(c => v.indexOf(norm_(c.nome)) >= 0 || norm_(c.nome).indexOf(v) >= 0)) || null;
}

function abaRespostas_() {
  const ss = ss_();
  const nome = cfg_().ABA_RESPOSTAS;
  if (nome) return ss.getSheetByName(nome);
  return ss.getSheets().find(s => /^(respostas ao formulário|form responses|respuestas de formulario)/i.test(s.getName())) || null;
}

function shTarefas_() { return ss_().getSheetByName(ABA.TAREFAS); }

function lerTarefas_() {
  const sh = shTarefas_();
  if (!sh || sh.getLastRow() < 2) return [];
  const valores = sh.getRange(2, 1, sh.getLastRow() - 1, CAMPOS.length).getDisplayValues();
  return valores.map((r, i) => {
    const t = { _linha: i + 2 };
    CAMPOS.forEach((c, j) => {
      let v = r[j];
      if (CAMPOS_JSON.includes(c[0])) { try { v = v ? JSON.parse(v) : []; } catch (e) { v = []; } }
      t[c[0]] = v;
    });
    return t;
  }).filter(t => t.id);
}

function buscarTarefa_(id) { return lerTarefas_().find(t => t.id === id) || null; }

function linhaDe_(t) {
  return [CAMPOS.map(c => {
    const v = t[c[0]];
    if (CAMPOS_JSON.includes(c[0])) return JSON.stringify(v || []);
    return v === undefined || v === null ? '' : String(v);
  })];
}

function inserirTarefa_(t) {
  const sh = shTarefas_();
  const linha = sh.getLastRow() + 1;
  sh.getRange(linha, 1, 1, CAMPOS.length).setNumberFormat('@').setValues(linhaDe_(t));
  t._linha = linha;
}

function salvarTarefa_(t) {
  if (!t._linha) { const x = buscarTarefa_(t.id); t._linha = x._linha; }
  shTarefas_().getRange(t._linha, 1, 1, CAMPOS.length).setNumberFormat('@').setValues(linhaDe_(t));
}

function registrar_(id, autor, acao, detalhe) {
  ss_().getSheetByName(ABA.HISTORICO).appendRow([agora_(), id, autor || '—', acao, detalhe || '']);
}

function contarComentarios_() {
  const o = {};
  lerAba_(ABA.COMENTARIOS).forEach(r => o[r[1]] = (o[r[1]] || 0) + 1);
  return o;
}

function toClient_(t) {
  const o = {};
  CAMPOS.forEach(c => o[c[0]] = t[c[0]]);
  return o;
}

/* ============================================================
 *  UTILITÁRIOS
 * ============================================================ */

function tz_() { return Session.getScriptTimeZone() || 'America/Campo_Grande'; }
function agora_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd HH:mm'); }
function hoje_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function uid_() { return Utilities.getUuid().slice(0, 8); }
function lista_(s, sep) { return String(s || '').split(sep || '|').map(x => x.trim()).filter(String); }
function semAcento_(s) { return String(s).normalize('NFD').replace(/[̀-ͯ]/g, ''); }
function norm_(s) { return semAcento_(String(s || '').trim().toLowerCase()); }
function escRe_(s) { return String(s).replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
function esc_(s) { return String(s || '').replace(/[&<>"]/g, ch => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[ch])); }

function somarDias_(iso, dias) {
  const [a, m, d] = iso.split('-').map(Number);
  const dt = new Date(a, m - 1, d + Number(dias));
  return Utilities.formatDate(dt, tz_(), 'yyyy-MM-dd');
}

function normalizarData_(s) {
  s = String(s).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (m) return `${m[1]}-${pad_(m[2])}-${pad_(m[3])}`;
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{2,4})/);
  if (m) return `${m[3].length === 2 ? '20' + m[3] : m[3]}-${pad_(m[2])}-${pad_(m[1])}`;
  const d = new Date(s);
  return isNaN(d) ? '' : Utilities.formatDate(d, tz_(), 'yyyy-MM-dd');
}

function pad_(n) { return String(n).padStart(2, '0'); }
function dataBr_(iso) { if (!iso) return '—'; const [a, m, d] = iso.slice(0, 10).split('-'); return `${d}/${m}/${a}`; }


/*************************************************
 PAINEL SEGURO - ARENA AQUÁTICA
 INSPEÇÃO E FREQUÊNCIA CONTROLADA
 BASES FIXAS:
 - Listageral
 - frequencia

 SISTEMA INDEPENDENTE VIA ID DA PLANILHA
*************************************************/

/* ================= CONFIGURAÇÃO ================= */

const ID_PLANILHA = 'COLOQUE_AQUI_O_ID_DA_PLANILHA';

const ABA_LISTAGERAL    = 'Listageral';
const ABA_FREQUENCIA    = 'frequencia';
const ABA_OPERADORES    = 'OPERADORES';
const ABA_LOG           = 'LOG_OPERADORES';

const TEMPO_SESSAO_HORAS = 6;
const CACHE_FILTROS_SEGUNDOS = 180;
const PAGE_SIZE_DEFAULT = 60;
const LIMITE_CONSULTA_PADRAO = PAGE_SIZE_DEFAULT;
const LIMITE_CONSULTA_MAXIMO = 150;

/* ================= WEB APP ================= */

function doGet() {
  return HtmlService
    .createHtmlOutputFromFile('index')
    .setTitle('Painel Seguro - Arena Aquática')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* ================= INSTALAÇÃO ================= */


function garantirEstruturaOperadores() {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  let abaOperadores = ss.getSheetByName(ABA_OPERADORES);
  if (!abaOperadores) {
    abaOperadores = ss.insertSheet(ABA_OPERADORES);
  }

  const cabecalhoEsperado = [
    'Matrícula',
    'Nome',
    'Senha',
    'Perfil',
    'Status',
    'TROCAR_SENHA',
    'Usuario',
    'SenhaHash',
    'Ultimo acesso',
    'Criado em',
    'Observacao',
    'CodigoServidor'
  ];

  if (abaOperadores.getLastRow() === 0) {
    abaOperadores.getRange(1, 1, 1, cabecalhoEsperado.length).setValues([cabecalhoEsperado]);
    abaOperadores.setFrozenRows(1);
    return;
  }

  const atual = abaOperadores.getRange(1, 1, 1, Math.max(abaOperadores.getLastColumn(), cabecalhoEsperado.length)).getDisplayValues()[0];
  const atualNorm = atual.map(normalizar);
  const faltantes = cabecalhoEsperado.filter(col => atualNorm.indexOf(normalizar(col)) === -1);

  if (faltantes.length > 0) {
    const inicio = abaOperadores.getLastColumn() + 1;
    abaOperadores.getRange(1, inicio, 1, faltantes.length).setValues([faltantes]);
  }

  abaOperadores.setFrozenRows(1);
}


function instalarSistemaArena() {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);

  let abaOperadores = ss.getSheetByName(ABA_OPERADORES);
  if (!abaOperadores) {
    abaOperadores = ss.insertSheet(ABA_OPERADORES);
  }

  let abaLog = ss.getSheetByName(ABA_LOG);
  if (!abaLog) {
    abaLog = ss.insertSheet(ABA_LOG);
  }

  const cabOperadores = [
    'Matricula',
    'Nome',
    'Usuario',
    'SenhaHash',
    'Perfil',
    'Status',
    'Ultimo acesso',
    'Criado em',
    'Observacao',
    'CodigoServidor'
  ];

  const cabLog = [
    'DataHora',
    'Usuario',
    'Nome',
    'Perfil',
    'Acao',
    'Detalhes',
    'AlunoID',
    'AlunoNome',
    'DataAula',
    'IP/Origem'
  ];

  prepararCabecalho(abaOperadores, cabOperadores);
  prepararCabecalho(abaLog, cabLog);

  garantirColunasFrequancia();

  const ultimaLinhaOperadores = abaOperadores.getLastRow();
  if (ultimaLinhaOperadores < 2) {
    const senhaInicial = '123456';
    const hash = gerarHashSenha(senhaInicial);

    abaOperadores.getRange(2, 1, 1, cabOperadores.length).setValues([[
      '0001',
      'Administrador Master',
      'admin',
      hash,
      'ADMIN',
      'ATIVO',
      '',
      new Date(),
      'Trocar a senha imediatamente após o primeiro acesso.',
      ''
    ]]);
  }

  SpreadsheetApp.flush();

  return {
    ok: true,
    mensagem: 'Sistema instalado com sucesso. Usuário inicial: admin | Senha inicial: 123456'
  };
}

/* ================= AUTENTICAÇÃO ================= */


function loginPainel(credenciais) {
  try {
    garantirEstruturaOperadores();

    const usuarioInformado = String((credenciais && credenciais.usuario) || '').trim();
    const senhaInformada = String((credenciais && credenciais.senha) || '').trim();

    if (!usuarioInformado || !senhaInformada) {
      throw new Error('Informe usuário/matrícula e senha.');
    }

    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName(ABA_OPERADORES);
    if (!aba || aba.getLastRow() < 2) {
      throw new Error('A aba OPERADORES não foi configurada.');
    }

    const dados = aba.getDataRange().getDisplayValues();
    const cab = dados[0];
    const mapa = mapearCabecalhos(cab);

    const idxMatricula = pegarIndiceObrigatorio(mapa, ['matricula']);
    const idxNome      = pegarIndiceObrigatorio(mapa, ['nome']);
    const idxSenha     = pegarIndiceOpcional(mapa, ['senha']);
    const idxSenhaHash = pegarIndiceOpcional(mapa, ['senhahash']);
    const idxPerfil    = pegarIndiceObrigatorio(mapa, ['perfil']);
    const idxStatus    = pegarIndiceObrigatorio(mapa, ['status']);
    const idxTrocar    = pegarIndiceOpcional(mapa, ['trocarsenha']);
    const idxUltimo    = pegarIndiceOpcional(mapa, ['ultimoacesso']);
    const idxCodServ   = pegarIndiceOpcional(mapa, ['codigoservidor']);
    const idxUsuario   = pegarIndiceOpcional(mapa, ['usuario']);

    const loginNorm = normalizar(usuarioInformado);

    for (let i = 1; i < dados.length; i++) {
      const row = dados[i];
      const matricula = String(row[idxMatricula] || '').trim();
      const nome = String(row[idxNome] || '').trim();
      const usuario = idxUsuario > -1 ? String(row[idxUsuario] || '').trim() : '';
      const senhaTexto = idxSenha > -1 ? String(row[idxSenha] || '').trim() : '';
      const senhaHash = idxSenhaHash > -1 ? String(row[idxSenhaHash] || '').trim() : '';
      const perfil = String(row[idxPerfil] || 'OPERADOR').trim().toUpperCase();
      const status = String(row[idxStatus] || 'ATIVO').trim().toUpperCase();
      const trocarSenha = idxTrocar > -1 ? String(row[idxTrocar] || '').trim().toUpperCase() : 'NÃO';
      const codigoServidor = idxCodServ > -1 ? String(row[idxCodServ] || '').trim() : '';

      const bateLogin =
        normalizar(matricula) === loginNorm ||
        (usuario && normalizar(usuario) === loginNorm);

      if (!bateLogin) continue;

      if (status !== 'ATIVO') {
        throw new Error('Seu acesso está inativo. Procure o administrador.');
      }

      let senhaValida = false;

      if (senhaHash) {
        senhaValida = senhaHash === gerarHashSenha(senhaInformada);
      } else if (senhaTexto) {
        senhaValida = senhaTexto === senhaInformada;
        if (senhaValida && idxSenhaHash > -1) {
          aba.getRange(i + 1, idxSenhaHash + 1).setValue(gerarHashSenha(senhaInformada));
        }
      }

      if (!senhaValida) {
        registrarLogInterno({
          usuario: usuario || matricula,
          nome: nome,
          perfil: perfil,
          acao: 'LOGIN_FALHOU',
          detalhes: 'Senha incorreta.'
        });
        throw new Error('Usuário ou senha inválidos.');
      }

      if (idxUltimo > -1) {
        aba.getRange(i + 1, idxUltimo + 1).setValue(new Date());
      }

      const token = Utilities.getUuid();
      const sessao = {
        token: token,
        rowOperador: i + 1,
        matricula: matricula,
        usuario: usuario || matricula,
        nome: nome,
        perfil: perfil,
        codigoServidor: codigoServidor,
        trocarSenha: trocarSenha === 'SIM',
        criadoEm: new Date().toISOString()
      };

      CacheService.getScriptCache().put(
        'sessao_' + token,
        JSON.stringify(sessao),
        TEMPO_SESSAO_HORAS * 3600
      );

      registrarLogInterno({
        usuario: sessao.usuario,
        nome: sessao.nome,
        perfil: sessao.perfil,
        acao: 'LOGIN_OK',
        detalhes: trocarSenha === 'SIM' ? 'Login realizado - troca obrigatória de senha.' : 'Login realizado com sucesso.'
      });

      return {
        ok: true,
        token: token,
        trocaSenhaObrigatoria: trocarSenha === 'SIM',
        usuario: {
          matricula: matricula,
          usuario: sessao.usuario,
          nome: nome,
          perfil: perfil
        }
      };
    }

    throw new Error('Usuário ou senha inválidos.');
  } catch (erro) {
    return { ok: false, mensagem: erro.message };
  }
}


function logoutPainel(token) {
  const sessao = validarSessao(token);
  CacheService.getScriptCache().remove('sessao_' + token);

  registrarLogInterno({
    usuario: sessao.usuario,
    nome: sessao.nome,
    perfil: sessao.perfil,
    acao: 'LOGOUT',
    detalhes: 'Logout realizado.'
  });

  return { ok: true };
}

function validarSessao(token) {
  const chave = 'sessao_' + String(token || '').trim();
  const bruto = CacheService.getScriptCache().get(chave);

  if (!bruto) {
    throw new Error('Sessão expirada. Entre novamente.');
  }

  return JSON.parse(bruto);
}


function obterBaseListageralCache() {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const aba = ss.getSheetByName(ABA_LISTAGERAL);

  if (!aba || aba.getLastRow() < 2) {
    throw new Error('A aba Listageral não possui dados.');
  }

  const dados = obterDadosAba(aba);
  return {
    cabecalho: dados[0],
    linhas: dados.slice(1)
  };
}

function limparCacheBaseListageral() {
  return { ok: true };
}

function obterDadosAba(aba) {
  if (!aba) return [];

  const ultimaLinha = aba.getLastRow();
  const ultimaColuna = aba.getLastColumn();

  if (ultimaLinha < 1 || ultimaColuna < 1) {
    return [];
  }

  return aba.getRange(1, 1, ultimaLinha, ultimaColuna).getDisplayValues();
}

function obterDadosTabulares(fonte) {
  if (!fonte) return [];
  return Array.isArray(fonte) ? fonte : obterDadosAba(fonte);
}


/* ================= FILTROS ================= */

function obterFiltrosPainel(token) {
  try {
    const sessao = validarSessao(token);
    const cache = CacheService.getScriptCache();
    const cacheIdentidade = normalizar(
      sessao.perfil === 'ADMIN'
        ? 'admin'
        : (sessao.codigoServidor || sessao.matricula || sessao.usuario || 'geral')
    );
    const cacheKey = 'filtros_painel_v1_' + cacheIdentidade;
    const cached = cache.get(cacheKey);
    if (cached) {
      return { ok: true, filtros: JSON.parse(cached) };
    }

    const base = obterBaseListageralCache();
    const filtros = montarFiltrosPainelComBase(sessao, base);

    cache.put(cacheKey, JSON.stringify(filtros), CACHE_FILTROS_SEGUNDOS);
    return { ok: true, filtros: filtros };
  } catch (erro) {
    registrarLogInterno({
      usuario: 'SISTEMA',
      nome: 'SISTEMA',
      perfil: 'ERRO',
      acao: 'ERRO_FILTROS',
      detalhes: valorSeguro(erro && erro.message)
    });
    return { ok: false, mensagem: erro.message || 'Falha ao carregar filtros.' };
  }
}

function montarFiltrosPainelComBase(sessao, base) {
  const cab = base.cabecalho;
  const dados = base.linhas;
  const mapa = mapearCabecalhos(cab);

  const idxTurno      = pegarIndiceOpcional(mapa, ['turno']);
  const idxHorario    = pegarIndiceOpcional(mapa, ['horario']);
  const idxModalidade = pegarIndiceOpcional(mapa, ['modalidade']);
  const idxDias       = pegarIndiceOpcional(mapa, ['dias']);
  const idxServidor   = pegarIndiceOpcional(mapa, ['servidor', 'codigoservidor']);
  const idxMatricula  = pegarIndiceOpcional(mapa, ['matricula']);

  const turnos = new Set();
  const horarios = new Set();
  const modalidades = new Set();
  const dias = new Set();
  const matriculas = new Set();

  for (let i = 0; i < dados.length; i++) {
    const row = dados[i];

    if (!podeVerLinha(sessao, row, {
      idxServidor: idxServidor,
      idxMatricula: idxMatricula
    })) {
      continue;
    }

    addSet(turnos, idxTurno > -1 ? row[idxTurno] : '');
    addSet(horarios, idxHorario > -1 ? row[idxHorario] : '');
    addSet(modalidades, idxModalidade > -1 ? row[idxModalidade] : '');
    addSet(dias, idxDias > -1 ? row[idxDias] : '');
    addSet(matriculas, idxMatricula > -1 ? row[idxMatricula] : '');
  }

  return {
    turnos: ordenarLista(Array.from(turnos).filter(Boolean)),
    horarios: ordenarLista(Array.from(horarios).filter(Boolean)),
    modalidades: ordenarLista(Array.from(modalidades).filter(Boolean)),
    dias: ordenarLista(Array.from(dias).filter(Boolean)),
    servidores: [],
    matriculas: ordenarLista(Array.from(matriculas).filter(Boolean))
  };
}

/* ================= BUSCA PAINEL ================= */


function buscarPainel(token, filtros) {
  try {
    const sessao = validarSessao(token);
    filtros = filtros || {};

    const pagina = Math.max(1, Number(filtros.pagina || 1));
    const limiteInformado = filtros.limite || filtros.pageSize || LIMITE_CONSULTA_PADRAO;
    const limite = Math.min(
      LIMITE_CONSULTA_MAXIMO,
      Math.max(1, Number(limiteInformado))
    );

    const dataSelecionada = normalizarDataEntrada((filtros && filtros.dataAula) || new Date());

    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const abaBase = ss.getSheetByName(ABA_LISTAGERAL);
    const abaFreq = ss.getSheetByName(ABA_FREQUENCIA);

    if (!abaBase || abaBase.getLastRow() < 2) {
      throw new Error('A aba Listageral não possui dados.');
    }

    const base = obterDadosAba(abaBase);
    const cabBase = base[0];
    const mapaBase = mapearCabecalhos(cabBase);

    const idxId         = pegarIndiceOpcional(mapaBase, ['id', 'codigo', 'idaluno']);
    const idxNome       = pegarIndiceOpcional(mapaBase, ['nomedoaluno', 'aluno', 'nome']);
    const idxRGCPF      = pegarIndiceOpcional(mapaBase, ['rgcpf', 'rg', 'cpf']);
    const idxEmail      = pegarIndiceOpcional(mapaBase, ['email', 'emailresponsavel', 'e-mail']);
    const idxModalidade = pegarIndiceOpcional(mapaBase, ['modalidade']);
    const idxTurno      = pegarIndiceOpcional(mapaBase, ['turno']);
    const idxHorario    = pegarIndiceOpcional(mapaBase, ['horario']);
    const idxDias       = pegarIndiceOpcional(mapaBase, ['dias']);
    const idxServidor   = pegarIndiceOpcional(mapaBase, ['servidor', 'codigoservidor']);
    const idxMatricula  = pegarIndiceOpcional(mapaBase, ['matricula']);
    const idxFoto       = pegarIndiceOpcional(mapaBase, ['foto2', 'foto']);

    if (idxNome === -1) {
      throw new Error('Não encontrei a coluna do nome do aluno na aba Listageral. Verifique o cabeçalho real da planilha.');
    }

    const freqDados = obterDadosTabulares(abaFreq);
    const freqMap = construirMapaFrequencia(freqDados, dataSelecionada);
    const inativosPermanentes = construirMapaInativosPermanentes(freqDados);
    const inicio = (pagina - 1) * limite;
    const fim = inicio + limite;
    let total = 0;
    const registrosPagina = [];

    for (let i = 1; i < base.length; i++) {
      const row = base[i];

      if (!podeVerLinha(sessao, row, {
        idxServidor: idxServidor,
        idxMatricula: idxMatricula
      })) {
        continue;
      }

      const item = {
        rowBase: i + 1,
        alunoId: idxId > -1 ? valorSeguro(row[idxId]) : '',
        nome: valorSeguro(row[idxNome]),
        rgCpf: idxRGCPF > -1 ? valorSeguro(row[idxRGCPF]) : '',
        email: idxEmail > -1 ? valorSeguro(row[idxEmail]) : '',
        modalidade: idxModalidade > -1 ? valorSeguro(row[idxModalidade]) : '',
        turno: idxTurno > -1 ? valorSeguro(row[idxTurno]) : '',
        horario: idxHorario > -1 ? valorSeguro(row[idxHorario]) : '',
        dias: idxDias > -1 ? valorSeguro(row[idxDias]) : '',
  
        matricula: idxMatricula > -1 ? valorSeguro(row[idxMatricula]) : '',
        foto: idxFoto > -1 ? transformarLinkFoto(row[idxFoto]) : ''
      };

      if (!passaFiltros(item, filtros)) continue;

      const chavesFreq = gerarChaveFrequenciaFlex(dataSelecionada, item.alunoId, item.nome, item.turno, item.horario, item.rgCpf);
      const freq =
        freqMap[chavesFreq.completa] ||
        freqMap[chavesFreq.porId] ||
        freqMap[chavesFreq.porRgCpf] ||
        freqMap[chavesFreq.porNomeTurno] ||
        freqMap[chavesFreq.porNome] ||
        {};

      const chavePermanente = obterChaveAlunoResumo({
        trId: item.alunoId,
        nome: item.nome,
        rgCpf: item.rgCpf,
        modalidade: item.modalidade,
        turno: item.turno,
        horario: item.horario,
        dias: item.dias
      });
      const permanente = inativosPermanentes[chavePermanente] || null;

      item.dataAula = dataSelecionada;
      item.rawStatusFrequencia = freq.rawStatusFrequencia || '';
      item.statusFrequencia = freq.statusFrequencia || ((permanente && permanente.status === 'INATIVO') ? 'INATIVO' : 'PENDENTE');
      item.observacao = freq.observacao || '';
      item.lancadoPor = freq.lancadoPor || '';
      item.ultimaAtualizacao = freq.ultimaAtualizacao || '';
      item.rowFrequencia = freq.rowFrequencia || '';

      total++;

      if (total > inicio && total <= fim) {
        registrosPagina.push(montarRegistroPainel(item));
      }
    }

    registrarLogInterno({
      usuario: sessao.usuario,
      nome: sessao.nome,
      perfil: sessao.perfil,
      acao: 'CONSULTA_PAINEL',
      detalhes: 'Consulta realizada em ' + dataSelecionada + ' | Total filtrado: ' + total + ' | Página: ' + pagina
    });

    const resumoDireto = gerarResumoDaFrequencia(freqDados, dataSelecionada, filtros, total);

    return {
      ok: true,
      dataAula: dataSelecionada,
      total: total,
      pagina: pagina,
      limite: limite,
      totalPaginas: Math.max(1, Math.ceil(total / limite)),
      registros: registrosPagina,
      resumo: resumoDireto
    };
  } catch (erro) {
    registrarLogInterno({
      usuario: 'SISTEMA',
      nome: 'SISTEMA',
      perfil: 'ERRO',
      acao: 'ERRO_BUSCA_PAINEL',
      detalhes: valorSeguro(erro && erro.message)
    });
    return { ok: false, mensagem: erro.message || 'Falha ao consultar o painel.' };
  }
}


/* ================= GRAVAÇÃO DE FREQUÊNCIA ================= */

function montarRegistroPainel(item) {
  return {
    rowBase: item.rowBase,
    alunoId: item.alunoId,
    nome: item.nome,
    rgCpf: item.rgCpf,
    email: item.email,
    modalidade: item.modalidade,
    turno: item.turno,
    horario: item.horario,
    dias: item.dias,
    servidor: item.servidor,
    matricula: item.matricula,
    foto: item.foto,
    dataAula: item.dataAula,
    statusFrequencia: item.statusFrequencia,
    observacao: item.observacao,
    lancadoPor: item.lancadoPor,
    ultimaAtualizacao: item.ultimaAtualizacao,
    rowFrequencia: item.rowFrequencia
  };
}

function salvarFrequenciaPainel(token, payload) {
  const lock = LockService.getScriptLock();

  try {
    lock.waitLock(20000);

    const sessao = validarSessao(token);
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const abaFreq = ss.getSheetByName(ABA_FREQUENCIA);

    if (!abaFreq) {
      throw new Error('A aba frequencia não foi encontrada.');
    }

    garantirColunasFrequancia();

    payload = payload || {};

    const dados = obterDadosAba(abaFreq);
    const cab = dados[0];
    const mapa = mapearCabecalhos(cab);

    const idxDataHora   = pegarIndiceObrigatorio(mapa, ['datahora']);
    const idxTrId       = pegarIndiceOpcional(mapa, ['trid', 'id']);
    const idxNome       = pegarIndiceObrigatorio(mapa, ['nome']);
    const idxRGCPF      = pegarIndiceOpcional(mapa, ['rgcpf', 'rg', 'cpf']);
    const idxModalidade = pegarIndiceObrigatorio(mapa, ['modalidade']);
    const idxTurno      = pegarIndiceObrigatorio(mapa, ['turno']);
    const idxHorario    = pegarIndiceObrigatorio(mapa, ['horario']);
    const idxDias       = pegarIndiceObrigatorio(mapa, ['dias']);
    const idxStatus     = pegarIndiceObrigatorio(mapa, ['presenteausente']);
    const idxJustific   = pegarIndiceObrigatorio(mapa, ['justificarausencia']);
    const idxPcd        = pegarIndiceOpcional(mapa, ['epcd', 'pcd']);
    const idxCid        = pegarIndiceOpcional(mapa, ['sesiminformeocid', 'cid']);
    const idxRespons    = pegarIndiceOpcional(mapa, ['nomedoresponsavel', 'responsavel']);
    const idxCel        = pegarIndiceOpcional(mapa, ['celwhatsapp', 'celular', 'whatsapp']);
    const idxServidor   = pegarIndiceObrigatorio(mapa, ['servidor']);

    const dataAula = normalizarDataEntrada(payload.dataAula || new Date());
    const agora = new Date();
    const dataHoraTexto = Utilities.formatDate(agora, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm:ss');

    const alunoId = valorSeguro(payload.alunoId);
    const alunoNome = valorSeguro(payload.nome);
    const rgCpf = valorSeguro(payload.rgCpf);
    const modalidade = valorSeguro(payload.modalidade);
    const turno = valorSeguro(payload.turno);
    const horario = valorSeguro(payload.horario);
    const dias = valorSeguro(payload.dias);
    const servidor = valorSeguro(sessao.codigoServidor || sessao.usuario);
    const observacao = valorSeguro(payload.observacao);
    const pcd = valorSeguro(payload.pcd);
    const cid = valorSeguro(payload.cid);
    const responsavel = valorSeguro(payload.nomeResponsavel);
    const celular = valorSeguro(payload.celWhatsapp);

    let statusEntrada = normalizarStatusFrequencia(payload.statusFrequencia);
    let statusPlanilha = statusEntrada;
    let justificativaPlanilha = '';

    if (statusEntrada === 'AUSENCIA JUSTIFICADA') {
      statusPlanilha = 'Ausente';
      justificativaPlanilha = observacao || 'Ausência justificada';
    } else if (statusEntrada === 'AUSENTE') {
      statusPlanilha = 'Ausente';
      justificativaPlanilha = observacao || '';
    } else if (statusEntrada === 'INATIVO') {
      statusPlanilha = 'Inativo';
      justificativaPlanilha = observacao || '';
    } else if (statusEntrada === 'ATIVO') {
      statusPlanilha = 'Ativo';
      justificativaPlanilha = observacao || '';
    } else {
      statusPlanilha = 'Presente';
      justificativaPlanilha = '';
    }

    if (!alunoNome) throw new Error('Nome do aluno não informado.');

    let linhaExistente = -1;
    for (let i = 1; i < dados.length; i++) {
      const row = dados[i];
      const dataLinha = extrairDataDeDataHora(row[idxDataHora] || '');
      const nomeLinha = valorSeguro(row[idxNome]);
      const turnoLinha = idxTurno > -1 ? valorSeguro(row[idxTurno]) : '';
      const horarioLinha = idxHorario > -1 ? valorSeguro(row[idxHorario]) : '';
      const rgLinha = idxRGCPF > -1 ? valorSeguro(row[idxRGCPF]) : '';
      const tridLinha = idxTrId > -1 ? valorSeguro(row[idxTrId]) : '';

      const mesmaData = dataLinha === dataAula;
      const mesmoAluno =
        (tridLinha && alunoId && normalizar(tridLinha) === normalizar(alunoId)) ||
        (rgLinha && rgCpf && normalizar(rgLinha) === normalizar(rgCpf)) ||
        normalizar(nomeLinha) === normalizar(alunoNome);

      const mesmaTurma =
        normalizar(turnoLinha) === normalizar(turno) &&
        normalizar(horarioLinha) === normalizar(horario);

      if (mesmaData && mesmoAluno && mesmaTurma) {
        linhaExistente = i + 1;
        break;
      }
    }

    const totalCols = abaFreq.getLastColumn();

    if (linhaExistente > -1) {
      const valores = abaFreq.getRange(linhaExistente, 1, 1, totalCols).getDisplayValues()[0];

      valores[idxDataHora] = dataHoraTexto;
      if (idxTrId > -1) valores[idxTrId] = alunoId;
      valores[idxNome] = alunoNome;
      if (idxRGCPF > -1) valores[idxRGCPF] = rgCpf;
      valores[idxModalidade] = modalidade;
      valores[idxTurno] = turno;
      valores[idxHorario] = horario;
      valores[idxDias] = dias;
      valores[idxStatus] = statusPlanilha;
      valores[idxJustific] = justificativaPlanilha;
      if (idxPcd > -1) valores[idxPcd] = pcd;
      if (idxCid > -1) valores[idxCid] = cid;
      if (idxRespons > -1) valores[idxRespons] = responsavel;
      if (idxCel > -1) valores[idxCel] = celular;
      valores[idxServidor] = servidor;

      abaFreq.getRange(linhaExistente, 1, 1, totalCols).setValues([valores]);
    } else {
      const novaLinha = new Array(totalCols).fill('');
      novaLinha[idxDataHora] = dataHoraTexto;
      if (idxTrId > -1) novaLinha[idxTrId] = alunoId;
      novaLinha[idxNome] = alunoNome;
      if (idxRGCPF > -1) novaLinha[idxRGCPF] = rgCpf;
      novaLinha[idxModalidade] = modalidade;
      novaLinha[idxTurno] = turno;
      novaLinha[idxHorario] = horario;
      novaLinha[idxDias] = dias;
      novaLinha[idxStatus] = statusPlanilha;
      novaLinha[idxJustific] = justificativaPlanilha;
      if (idxPcd > -1) novaLinha[idxPcd] = pcd;
      if (idxCid > -1) novaLinha[idxCid] = cid;
      if (idxRespons > -1) novaLinha[idxRespons] = responsavel;
      if (idxCel > -1) novaLinha[idxCel] = celular;
      novaLinha[idxServidor] = servidor;

      abaFreq.appendRow(novaLinha);
    }

    SpreadsheetApp.flush();

    registrarLogInterno({
      usuario: sessao.usuario,
      nome: sessao.nome,
      perfil: sessao.perfil,
      acao: 'SALVOU_FREQUENCIA',
      detalhes: 'Status: ' + statusEntrada,
      alunoId: alunoId,
      alunoNome: alunoNome,
      dataAula: dataAula
    });

    return { ok: true, mensagem: 'Frequência salva com sucesso.' };
  } catch (erro) {
    return { ok: false, mensagem: erro.message || 'Falha ao salvar frequência.' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* ================= LEITURA DE AUDITORIA ================= */


function obterDetalheAluno(token, payload) {
  try {
    const sessao = validarSessao(token);
    payload = payload || {};

    const rowBase = Number(payload.rowBase || 0);
    const dataAula = normalizarDataEntrada(payload.dataAula || new Date());

    if (!rowBase || rowBase < 2) {
      throw new Error('Linha do aluno inválida.');
    }

    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const abaBase = ss.getSheetByName(ABA_LISTAGERAL);
    const abaFreq = ss.getSheetByName(ABA_FREQUENCIA);

    if (!abaBase || abaBase.getLastRow() < rowBase) {
      throw new Error('Aluno não encontrado na Listageral.');
    }

    const cabecalho = abaBase.getRange(1, 1, 1, abaBase.getLastColumn()).getDisplayValues()[0];
    const linha = abaBase.getRange(rowBase, 1, 1, abaBase.getLastColumn()).getDisplayValues()[0];
    const mapa = mapearCabecalhos(cabecalho);

    const idxNome       = pegarIndiceOpcional(mapa, ['nomedoaluno', 'aluno', 'nome']);
    const idxId         = pegarIndiceOpcional(mapa, ['id', 'codigo', 'idaluno']);
    const idxRGCPF      = pegarIndiceOpcional(mapa, ['rgcpf', 'rg', 'cpf']);
    const idxEmail      = pegarIndiceOpcional(mapa, ['email', 'emailresponsavel', 'e-mail']);
    const idxModalidade = pegarIndiceOpcional(mapa, ['modalidade']);
    const idxTurno      = pegarIndiceOpcional(mapa, ['turno']);
    const idxHorario    = pegarIndiceOpcional(mapa, ['horario']);
    const idxDias       = pegarIndiceOpcional(mapa, ['dias']);
    const idxServidor   = pegarIndiceOpcional(mapa, ['servidor', 'codigoservidor']);
    const idxMatricula  = pegarIndiceOpcional(mapa, ['matricula']);
    const idxFoto       = pegarIndiceOpcional(mapa, ['foto2', 'foto']);
    const idxCracha     = pegarIndiceOpcional(mapa, ['mergeddocurlcracha2', 'linktomergeddoccracha2', 'cracha', 'crachapdf', 'linkcracha']);
    const idxPcd        = pegarIndiceOpcional(mapa, ['epcd', 'pcd']);
    const idxCid        = pegarIndiceOpcional(mapa, ['sesiminformeocid', 'cid']);
    const idxRespons    = pegarIndiceOpcional(mapa, ['nomedoresponsavel', 'responsavel']);
    const idxCel        = pegarIndiceOpcional(mapa, ['celwhatsapp', 'celular', 'whatsapp']);

    if (!podeVerLinha(sessao, linha, {
      idxServidor: idxServidor,
      idxMatricula: idxMatricula
    })) {
      throw new Error('Você não tem permissão para visualizar este aluno.');
    }

    const nome = idxNome > -1 ? valorSeguro(linha[idxNome]) : '';
    const alunoId = idxId > -1 ? valorSeguro(linha[idxId]) : '';
    const turno = idxTurno > -1 ? valorSeguro(linha[idxTurno]) : '';
    const horario = idxHorario > -1 ? valorSeguro(linha[idxHorario]) : '';

    const freqMap = construirMapaFrequencia(abaFreq, dataAula);
    const chavesFreq = gerarChaveFrequenciaFlex(dataAula, alunoId, nome, turno, horario, idxRGCPF > -1 ? valorSeguro(linha[idxRGCPF]) : '');
    const freq =
      freqMap[chavesFreq.completa] ||
      freqMap[chavesFreq.porRgCpf] ||
      freqMap[chavesFreq.porId] ||
      freqMap[chavesFreq.porNomeTurno] ||
      freqMap[chavesFreq.porNome] ||
      {};

    const campos = [];
    for (let i = 0; i < cabecalho.length; i++) {
      const titulo = valorSeguro(cabecalho[i]);
      const valor = valorSeguro(linha[i]);
      if (titulo) {
        campos.push({ campo: titulo, valor: valor });
      }
    }

    registrarLogInterno({
      usuario: sessao.usuario,
      nome: sessao.nome,
      perfil: sessao.perfil,
      acao: 'ABRIU_FICHA_ALUNO',
      detalhes: 'Ficha aberta no painel.',
      alunoId: alunoId,
      alunoNome: nome,
      dataAula: dataAula
    });

    const resumoHistorico = calcularResumoAlunoNaFrequencia(ss, {
      alunoId: alunoId,
      nome: nome,
      rgCpf: idxRGCPF > -1 ? valorSeguro(linha[idxRGCPF]) : ''
    });

    return {
      ok: true,
      aluno: {
        rowBase: rowBase,
        alunoId: alunoId,
        nome: nome,
        rgCpf: idxRGCPF > -1 ? valorSeguro(linha[idxRGCPF]) : '',
        email: idxEmail > -1 ? valorSeguro(linha[idxEmail]) : '',
        modalidade: idxModalidade > -1 ? valorSeguro(linha[idxModalidade]) : '',
        turno: turno,
        horario: horario,
        dias: idxDias > -1 ? valorSeguro(linha[idxDias]) : '',

        matricula: idxMatricula > -1 ? valorSeguro(linha[idxMatricula]) : '',
        foto: idxFoto > -1 ? transformarLinkFoto(linha[idxFoto]) : '',
        cracha: idxCracha > -1 ? transformarLinkCracha(linha[idxCracha]) : '',
        pcd: idxPcd > -1 ? valorSeguro(linha[idxPcd]) : '',
        cid: idxCid > -1 ? valorSeguro(linha[idxCid]) : '',
        nomeResponsavel: idxRespons > -1 ? valorSeguro(linha[idxRespons]) : '',
        celWhatsapp: idxCel > -1 ? valorSeguro(linha[idxCel]) : '',
        dataAula: dataAula,
        rawStatusFrequencia: freq.rawStatusFrequencia || '',
        statusFrequencia: freq.statusFrequencia || 'PENDENTE',
        observacao: freq.observacao || '',
        resumoHistorico: resumoHistorico,
        campos: campos
      }
    };
  } catch (erro) {
    return { ok: false, mensagem: erro.message || 'Falha ao abrir a ficha do aluno.' };
  }
}



function calcularResumoAlunoNaFrequencia(ss, aluno) {
  const resumo = {
    presentes: 0,
    ausentes: 0,
    justificadas: 0,
    ativos: 0,
    inativos: 0,
    percentualPresenca: 0,
    totalLancamentos: 0,
    alertaFrequencia: ''
  };

  const abaFreq = ss.getSheetByName(ABA_FREQUENCIA);
  if (!abaFreq || abaFreq.getLastRow() < 2) return resumo;

  const dados = abaFreq.getDataRange().getDisplayValues();
  const cab = dados[0];
  const mapa = mapearCabecalhos(cab);

  const idxId       = pegarIndiceOpcional(mapa, ['id', 'trid']);
  const idxNome     = pegarIndiceOpcional(mapa, ['nome', 'nomedoaluno', 'aluno']);
  const idxRGCPF    = pegarIndiceOpcional(mapa, ['rgcpf', 'rg', 'cpf']);
  const idxStatus   = pegarIndiceObrigatorio(mapa, ['presenteausente']);

  for (let i = 1; i < dados.length; i++) {
    const row = dados[i];

    const idLinha = idxId > -1 ? valorSeguro(row[idxId]) : '';
    const nomeLinha = idxNome > -1 ? valorSeguro(row[idxNome]) : '';
    const rgCpfLinha = idxRGCPF > -1 ? valorSeguro(row[idxRGCPF]) : '';

    const mesmoAluno =
      (aluno.alunoId && idLinha && normalizar(aluno.alunoId) === normalizar(idLinha)) ||
      (aluno.rgCpf && rgCpfLinha && normalizar(aluno.rgCpf) === normalizar(rgCpfLinha)) ||
      (aluno.nome && nomeLinha && normalizar(aluno.nome) === normalizar(nomeLinha));

    if (!mesmoAluno) continue;

    const valorColunaI = normalizar(valorSeguro(row[idxStatus]));

    if (valorColunaI === 'presente') resumo.presentes++;
    else if (valorColunaI === 'ausente') resumo.ausentes++;
    else if (valorColunaI === 'ausenciajustificada') resumo.justificadas++;
    else if (valorColunaI === 'ativo') resumo.ativos++;
    else if (valorColunaI === 'inativo') resumo.inativos++;
  }

  const totalConsiderado = resumo.presentes + resumo.ausentes + resumo.justificadas;
  resumo.totalLancamentos = totalConsiderado;

  if (totalConsiderado > 0) {
    resumo.percentualPresenca = Math.round((resumo.presentes / totalConsiderado) * 100);
  } else {
    resumo.percentualPresenca = 0;
  }

  if (totalConsiderado >= 4 && resumo.percentualPresenca < 75) {
    resumo.alertaFrequencia = 'Atenção: frequência abaixo de 75%.';
  } else if (totalConsiderado >= 4 && resumo.percentualPresenca < 85) {
    resumo.alertaFrequencia = 'Acompanhamento: frequência abaixo de 85%.';
  } else {
    resumo.alertaFrequencia = '';
  }

  return resumo;
}


function listarAuditoria(token, limite) {
  try {
    const sessao = validarSessao(token);
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName(ABA_LOG);

    if (!aba || aba.getLastRow() < 2) {
      return { ok: true, registros: [] };
    }

    const max = Math.max(1, Math.min(200, Number(limite || 60)));
    const ultimaLinha = aba.getLastRow();
    const ultimaColuna = aba.getLastColumn();
    const inicio = Math.max(2, ultimaLinha - max + 1);

    const cab = aba.getRange(1, 1, 1, ultimaColuna).getDisplayValues()[0];
    const dados = aba.getRange(inicio, 1, ultimaLinha - inicio + 1, ultimaColuna).getDisplayValues();

    const lista = [];
    for (let i = dados.length - 1; i >= 0; i--) {
      const row = dados[i];
      const item = {};
      cab.forEach(function(col, idx) {
        item[col] = row[idx];
      });

      if (sessao.perfil !== 'ADMIN' && sessao.perfil !== 'SUPERVISOR') {
        if (normalizar(item.Usuario || '') !== normalizar(sessao.usuario || '')) {
          continue;
        }
      }

      lista.push(item);
    }

    return { ok: true, registros: lista };
  } catch (erro) {
    return { ok: false, mensagem: erro.message || 'Falha ao carregar auditoria.' };
  }
}

/* ================= FUNÇÕES INTERNAS ================= */

function garantirColunasFrequancia() {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const aba = ss.getSheetByName(ABA_FREQUENCIA);

  if (!aba) {
    throw new Error('A aba frequencia não existe na planilha.');
  }

  if (aba.getLastRow() < 1 || aba.getLastColumn() < 1) {
    throw new Error('A aba frequencia está sem cabeçalhos.');
  }

  const cabecalho = aba.getRange(1, 1, 1, aba.getLastColumn()).getDisplayValues()[0];
  const mapa = mapearCabecalhos(cabecalho);

  const obrigatorios = [
    ['datahora', 'datahora'],
    ['nome', 'nome'],
    ['modalidade', 'modalidade'],
    ['turno', 'turno'],
    ['horario', 'horario'],
    ['dias', 'dias'],
    ['presenteausente', 'presenteausente'],
    ['justificarausencia', 'justificarausencia'],
    ['servidor', 'servidor']
  ];

  const faltantes = obrigatorios
    .filter(function(item) { return pegarIndiceOpcional(mapa, [item[0]]) === -1; })
    .map(function(item) { return item[1]; });

  if (faltantes.length) {
    throw new Error('A aba frequencia não está no formato esperado. Cabeçalhos não encontrados: ' + faltantes.join(', '));
  }

  return true;
}



function construirMapaInativosPermanentes(abaFreq) {
  const mapa = {};
  const dados = obterDadosTabulares(abaFreq);
  if (dados.length < 2) return mapa;
  const cab = dados[0];
  const cabMap = mapearCabecalhos(cab);

  const idxDataHora   = pegarIndiceOpcional(cabMap, ['datahora']);
  const idxId         = pegarIndiceOpcional(cabMap, ['trid', 'id']);
  const idxNome       = pegarIndiceOpcional(cabMap, ['nome', 'nomedoaluno', 'aluno']);
  const idxRGCPF      = pegarIndiceOpcional(cabMap, ['rgcpf', 'rg', 'cpf']);
  const idxModalidade = pegarIndiceOpcional(cabMap, ['modalidade']);
  const idxTurno      = pegarIndiceOpcional(cabMap, ['turno']);
  const idxHorario    = pegarIndiceOpcional(cabMap, ['horario']);
  const idxDias       = pegarIndiceOpcional(cabMap, ['dias']);
  const idxStatus     = pegarIndiceOpcional(cabMap, ['presenteausente']);

  if (idxStatus === -1) return mapa;

  for (let i = 1; i < dados.length; i++) {
    const row = dados[i];
    const rowObj = {
      trId: idxId > -1 ? valorSeguro(row[idxId]) : '',
      nome: idxNome > -1 ? valorSeguro(row[idxNome]) : '',
      rgCpf: idxRGCPF > -1 ? valorSeguro(row[idxRGCPF]) : '',
      modalidade: idxModalidade > -1 ? valorSeguro(row[idxModalidade]) : '',
      turno: idxTurno > -1 ? valorSeguro(row[idxTurno]) : '',
      horario: idxHorario > -1 ? valorSeguro(row[idxHorario]) : '',
      dias: idxDias > -1 ? valorSeguro(row[idxDias]) : ''
    };
    const chave = obterChaveAlunoResumo(rowObj);
    const ordem = obterDataHoraOrdenavel(idxDataHora > -1 ? row[idxDataHora] : '');
    const classe = classificarStatusResumo(row[idxStatus]);

    if (!mapa[chave] || ordem >= mapa[chave].ordem) {
      mapa[chave] = { ordem: ordem, status: classe };
    }
  }

  return mapa;
}


function construirMapaFrequencia(abaFreq, dataAula) {
  const mapaResultado = {};
  const dados = obterDadosTabulares(abaFreq);
  if (dados.length < 2) return mapaResultado;
  const cab = dados[0];
  const mapa = mapearCabecalhos(cab);

  const idxDataHora     = pegarIndiceOpcional(mapa, ['datahora']);
  const idxTrId         = pegarIndiceOpcional(mapa, ['trid', 'id']);
  const idxAlunoNome    = pegarIndiceOpcional(mapa, ['nome', 'nomedoaluno', 'aluno']);
  const idxRGCPF        = pegarIndiceOpcional(mapa, ['rgcpf', 'rg', 'cpf']);
  const idxModalidade   = pegarIndiceOpcional(mapa, ['modalidade']);
  const idxTurno        = pegarIndiceOpcional(mapa, ['turno']);
  const idxHorario      = pegarIndiceOpcional(mapa, ['horario']);
  const idxDias         = pegarIndiceOpcional(mapa, ['dias']);
  const idxStatus       = pegarIndiceOpcional(mapa, ['presenteausente', 'statusfrequencia', 'frequencia', 'status']);
  const idxJustific     = pegarIndiceOpcional(mapa, ['justificarausencia', 'justificativa', 'observacao', 'obs']);
  const idxServidor     = pegarIndiceOpcional(mapa, ['servidor']);

  if (idxDataHora === -1 || idxAlunoNome === -1 || idxStatus === -1) return mapaResultado;

  for (let i = 1; i < dados.length; i++) {
    const row = dados[i];
    const dataLinha = extrairDataDeDataHora(row[idxDataHora] || '');
    if (dataLinha !== dataAula) continue;

    const alunoId = idxTrId > -1 ? row[idxTrId] : '';
    const alunoNome = idxAlunoNome > -1 ? row[idxAlunoNome] : '';
    const turno = idxTurno > -1 ? row[idxTurno] : '';
    const horario = idxHorario > -1 ? row[idxHorario] : '';
    const rgCpf = idxRGCPF > -1 ? row[idxRGCPF] : '';

    const rawStatus = valorSeguro(idxStatus > -1 ? row[idxStatus] : '');
    const statusNormalizado = normalizarStatusFrequencia(rawStatus);
    const justificativa = idxJustific > -1 ? valorSeguro(row[idxJustific]) : '';

    const chaves = gerarChaveFrequenciaFlex(dataLinha, alunoId, alunoNome, turno, horario, rgCpf);

    const payload = {
      rawStatusFrequencia: rawStatus,
      statusFrequencia: justificativa ? 'AUSENCIA JUSTIFICADA' : statusNormalizado,
      observacao: justificativa,
      lancadoPor: '',
      ultimaAtualizacao: idxDataHora > -1 ? valorSeguro(row[idxDataHora]) : '',
      rowFrequencia: i + 1,
      rgCpf: rgCpf,
      dias: idxDias > -1 ? valorSeguro(row[idxDias]) : ''
    };

    mapaResultado[chaves.completa] = payload;
    if (normalizar(alunoNome || '')) mapaResultado[chaves.porNome] = payload;
    if (normalizar(alunoId || '')) mapaResultado[chaves.porId] = payload;
    if (normalizar(alunoNome || '') || normalizar(turno || '')) mapaResultado[chaves.porNomeTurno] = payload;
    if (normalizar(rgCpf || '')) mapaResultado[chaves.porRgCpf] = payload;
  }

  return mapaResultado;
}


function passaFiltros(item, filtros) {
  filtros = filtros || {};

  const busca = normalizar(filtros.busca || '');

  if (filtros.turno && item.turno !== filtros.turno) return false;
  if (filtros.horario && item.horario !== filtros.horario) return false;
  if (filtros.modalidade && item.modalidade !== filtros.modalidade) return false;
  if (filtros.dias && item.dias !== filtros.dias) return false;
  if (filtros.servidor && item.servidor !== filtros.servidor) return false;
  if (filtros.matricula && String(item.matricula) !== String(filtros.matricula)) return false;

  if (busca) {
    const alvo = normalizar([
      item.nome,
      item.alunoId,
      item.rgCpf,
      item.email,
      item.modalidade
    ].join(' '));

    if (alvo.indexOf(busca) === -1) return false;
  }

  return true;
}


function linhaFrequenciaPassaFiltros(rowObj, filtros) {
  filtros = filtros || {};
  const busca = normalizar(filtros.busca || '');

  if (filtros.turno && valorSeguro(rowObj.turno) !== valorSeguro(filtros.turno)) return false;
  if (filtros.horario && valorSeguro(rowObj.horario) !== valorSeguro(filtros.horario)) return false;
  if (filtros.modalidade && valorSeguro(rowObj.modalidade) !== valorSeguro(filtros.modalidade)) return false;
  if (filtros.dias && valorSeguro(rowObj.dias) !== valorSeguro(filtros.dias)) return false;


  if (busca) {
    const alvo = normalizar([
      rowObj.nome,
      rowObj.trId,
      rowObj.rgCpf,
      rowObj.modalidade,
      rowObj.turno,
      rowObj.horario
    ].join(' '));

    if (alvo.indexOf(busca) === -1) return false;
  }

  return true;
}


function obterDataHoraOrdenavel(valor) {
  const txt = valorSeguro(valor);
  if (!txt) return 0;
  const parsed = new Date(txt.replace(/(\d{2})\/(\d{2})\/(\d{4})\s*(\d{2}:\d{2}:\d{2})?/, '$3-$2-$1 $4'));
  if (!isNaN(parsed.getTime())) return parsed.getTime();
  return 0;
}

function obterChaveAlunoResumo(rowObj) {
  return [
    normalizar(rowObj.trId || ''),
    normalizar(rowObj.rgCpf || ''),
    normalizar(rowObj.nome || ''),
    normalizar(rowObj.turno || ''),
    normalizar(rowObj.horario || ''),
    normalizar(rowObj.dias || ''),
    normalizar(rowObj.modalidade || '')
  ].join('|');
}

function classificarStatusResumo(valorColunaI) {
  const v = normalizar(valorSeguro(valorColunaI));
  if (v === 'presente') return 'PRESENTE';
  if (v === 'ausente') return 'AUSENTE';
  if (v === 'ausenciajustificada' || v === 'justificada' || v === 'justificado') return 'JUSTIFICADA';
  if (v === 'ativo') return 'ATIVO';
  if (v === 'inativo') return 'INATIVO';
  return '';
}


function gerarResumoDaFrequencia(abaFreq, dataAula, filtros, totalListageral) {
  const resumo = {
    total: Number(totalListageral || 0),
    presentes: 0,
    ausentes: 0,
    justificadas: 0,
    ativos: 0,
    inativos: 0,
    pendentes: 0
  };

  const dados = obterDadosTabulares(abaFreq);

  if (dados.length < 2) {
    resumo.pendentes = resumo.total;
    return resumo;
  }

  const cab = dados[0];
  const mapa = mapearCabecalhos(cab);

  const idxDataHora   = pegarIndiceOpcional(mapa, ['datahora']);
  const idxTrId       = pegarIndiceOpcional(mapa, ['trid', 'id']);
  const idxNome       = pegarIndiceOpcional(mapa, ['nome', 'nomedoaluno', 'aluno']);
  const idxRGCPF      = pegarIndiceOpcional(mapa, ['rgcpf', 'rg', 'cpf']);
  const idxModalidade = pegarIndiceOpcional(mapa, ['modalidade']);
  const idxTurno      = pegarIndiceOpcional(mapa, ['turno']);
  const idxHorario    = pegarIndiceOpcional(mapa, ['horario']);
  const idxDias       = pegarIndiceOpcional(mapa, ['dias']);
  const idxStatus     = pegarIndiceObrigatorio(mapa, ['presenteausente']);

  const ultimosDoDia = {};
  const ultimosGerais = {};

  for (let i = 1; i < dados.length; i++) {
    const row = dados[i];
    const rowObj = {
      trId: idxTrId > -1 ? valorSeguro(row[idxTrId]) : '',
      nome: idxNome > -1 ? valorSeguro(row[idxNome]) : '',
      rgCpf: idxRGCPF > -1 ? valorSeguro(row[idxRGCPF]) : '',
      modalidade: idxModalidade > -1 ? valorSeguro(row[idxModalidade]) : '',
      turno: idxTurno > -1 ? valorSeguro(row[idxTurno]) : '',
      horario: idxHorario > -1 ? valorSeguro(row[idxHorario]) : '',
      dias: idxDias > -1 ? valorSeguro(row[idxDias]) : '',
      status: idxStatus > -1 ? valorSeguro(row[idxStatus]) : '',
      dataHora: idxDataHora > -1 ? valorSeguro(row[idxDataHora]) : ''
    };

    if (!linhaFrequenciaPassaFiltros(rowObj, filtros)) continue;

    const chave = obterChaveAlunoResumo(rowObj);
    const ordem = obterDataHoraOrdenavel(rowObj.dataHora);
    const classe = classificarStatusResumo(rowObj.status);

    if (!ultimosGerais[chave] || ordem >= ultimosGerais[chave].ordem) {
      ultimosGerais[chave] = { ordem: ordem, status: classe };
    }

    const dataLinha = extrairDataDeDataHora(rowObj.dataHora || '');
    if (dataLinha !== dataAula) continue;

    if (!ultimosDoDia[chave] || ordem >= ultimosDoDia[chave].ordem) {
      ultimosDoDia[chave] = { ordem: ordem, status: classe };
    }
  }

  Object.keys(ultimosGerais).forEach(function(chave) {
    if (ultimosGerais[chave].status === 'INATIVO') {
      resumo.inativos++;
    }
  });

  Object.keys(ultimosDoDia).forEach(function(chave) {
    const classe = ultimosDoDia[chave].status;
    if (classe === 'PRESENTE') resumo.presentes++;
    else if (classe === 'AUSENTE') resumo.ausentes++;
    else if (classe === 'JUSTIFICADA') resumo.justificadas++;
    else if (classe === 'ATIVO') resumo.ativos++;
  });

  const contabilizados = resumo.presentes + resumo.ausentes + resumo.justificadas + resumo.ativos + resumo.inativos;
  resumo.pendentes = Math.max(resumo.total - contabilizados, 0);

  return resumo;
}


function gerarResumo(registros) {
  const resumo = {
    total: registros.length,
    presentes: 0,
    ausentes: 0,
    justificadas: 0,
    ativos: 0,
    inativos: 0,
    pendentes: 0
  };

  registros.forEach(function(item) {
    const raw = normalizarStatusFrequencia(item.rawStatusFrequencia || '');

    if (raw === 'PRESENTE') resumo.presentes++;
    else if (raw === 'AUSENTE') resumo.ausentes++;
    else if (raw === 'ATIVO') resumo.ativos++;
    else if (raw === 'INATIVO') resumo.inativos++;
    else resumo.pendentes++;

    if (valorSeguro(item.observacao || '') !== '') {
      resumo.justificadas++;
    }
  });

  return resumo;
}

function prepararCabecalho(aba, cabecalho) {
  if (aba.getLastRow() === 0) {
    aba.getRange(1, 1, 1, cabecalho.length).setValues([cabecalho]);
    aba.setFrozenRows(1);
    return;
  }

  const atual = aba.getRange(1, 1, 1, Math.max(aba.getLastColumn(), cabecalho.length)).getValues()[0];
  const atualNorm = atual.map(normalizar);

  const faltantes = cabecalho.filter(function(col) {
    return atualNorm.indexOf(normalizar(col)) === -1;
  });

  if (faltantes.length > 0) {
    const inicio = aba.getLastColumn() + 1;
    aba.getRange(1, inicio, 1, faltantes.length).setValues([faltantes]);
  }

  aba.setFrozenRows(1);
}

function registrarLogInterno(obj) {
  try {
    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    let aba = ss.getSheetByName(ABA_LOG);

    if (!aba) {
      aba = ss.insertSheet(ABA_LOG);
      prepararCabecalho(aba, [
        'DataHora',
        'Usuario',
        'Nome',
        'Perfil',
        'Acao',
        'Detalhes',
        'AlunoID',
        'AlunoNome',
        'DataAula',
        'IP/Origem'
      ]);
    }

    aba.appendRow([
      new Date(),
      valorSeguro(obj.usuario),
      valorSeguro(obj.nome),
      valorSeguro(obj.perfil),
      valorSeguro(obj.acao),
      valorSeguro(obj.detalhes),
      valorSeguro(obj.alunoId),
      valorSeguro(obj.alunoNome),
      valorSeguro(obj.dataAula),
      'WEBAPP'
    ]);
  } catch (e) {}
}

function podeVerLinha(sessao, row, indices) {
  // Regra atual: qualquer usuário autenticado pode visualizar e editar a ficha do aluno.
  return true;
}

function gerarChaveFrequencia(dataAula, alunoId, nome, turno, horario) {
  return [
    normalizarDataEntrada(dataAula),
    normalizar(alunoId || ''),
    normalizar(nome || ''),
    normalizar(turno || ''),
    normalizar(horario || '')
  ].join('|');
}

function mapearCabecalhos(cabecalho) {
  const mapa = {};
  cabecalho.forEach(function(nome, idx) {
    mapa[normalizar(nome)] = idx;
  });
  return mapa;
}

function pegarIndiceObrigatorio(mapa, aliases) {
  const idx = pegarIndiceOpcional(mapa, aliases);
  if (idx === -1) {
    throw new Error('Cabeçalho obrigatório não encontrado: ' + aliases.join(' / '));
  }
  return idx;
}

function pegarIndiceOpcional(mapa, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const chave = normalizar(aliases[i]);
    if (Object.prototype.hasOwnProperty.call(mapa, chave)) {
      return mapa[chave];
    }
  }
  return -1;
}

function addSet(setObj, valor) {
  const txt = valorSeguro(valor);
  if (txt) setObj.add(txt);
}

function ordenarLista(lista) {
  return lista
    .filter(function(v) { return String(v || '').trim() !== ''; })
    .sort(function(a, b) {
      return String(a).localeCompare(String(b), 'pt-BR', { sensitivity: 'base' });
    });
}

function valorSeguro(v) {
  return v === null || v === undefined ? '' : String(v).trim();
}

function normalizar(texto) {
  return String(texto || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[^a-zA-Z0-9]/g, '')
    .toLowerCase()
    .trim();
}

function normalizarDataEntrada(valor) {
  if (!valor) {
    return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }

  if (Object.prototype.toString.call(valor) === '[object Date]' && !isNaN(valor.getTime())) {
    return Utilities.formatDate(valor, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }

  const txt = String(valor).trim();

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(txt)) {
    return txt;
  }

  const d = new Date(txt);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  }

  return txt;
}



function transformarLinkCracha(url) {
  const txt = valorSeguro(url);
  if (!txt) return '';

  let match = txt.match(/\/d\/([a-zA-Z0-9_-]+)/);
  if (!match) {
    match = txt.match(/id=([a-zA-Z0-9_-]+)/);
  }

  if (match && match[1]) {
    return 'https://drive.google.com/file/d/' + match[1] + '/view';
  }

  return txt;
}

function transformarLinkFoto(url) {
  const txt = valorSeguro(url);
  if (!txt) return '';

  let match = txt.match(/id=([a-zA-Z0-9_-]+)/);
  if (!match) {
    match = txt.match(/\/d\/([a-zA-Z0-9_-]+)/);
  }

  if (match && match[1]) {
    return 'https://drive.google.com/thumbnail?id=' + match[1] + '&sz=w400-h520';
  }

  return txt;
}


function diagnosticarEstruturaArena() {
  const ss = SpreadsheetApp.openById(ID_PLANILHA);
  const nomesAbas = ss.getSheets().map(s => s.getName());

  const resultado = {
    planilha: ss.getName(),
    abas: nomesAbas,
    listageral: null,
    frequencia: null
  };

  const abaLista = ss.getSheetByName(ABA_LISTAGERAL);
  if (abaLista) {
    const ultCol = Math.max(1, abaLista.getLastColumn());
    resultado.listageral = abaLista.getRange(1, 1, 1, ultCol).getValues()[0];
  }

  const abaFreq = ss.getSheetByName(ABA_FREQUENCIA);
  if (abaFreq) {
    const ultColF = Math.max(1, abaFreq.getLastColumn());
    resultado.frequencia = abaFreq.getRange(1, 1, 1, ultColF).getValues()[0];
  }

  return resultado;
}


function normalizarStatusFrequencia(status) {
  const st = normalizar(valorSeguro(status));

  if (!st) return 'PENDENTE';
  if (st === 'presente') return 'PRESENTE';
  if (st === 'ausente') return 'AUSENTE';
  if (st === 'ativo') return 'ATIVO';
  if (st === 'inativo') return 'INATIVO';
  if (st === 'ausenciajustificada' || st === 'ausenciajustificativa' || st === 'justificada' || st === 'faltajustificada') {
    return 'AUSENCIA JUSTIFICADA';
  }

  return String(valorSeguro(status)).toUpperCase();
}

function gerarChaveFrequenciaFlex(dataAula, alunoId, nome, turno, horario, rgCpf) {
  return {
    completa: gerarChaveFrequencia(dataAula, alunoId, nome, turno, horario),
    porNome: [normalizarDataEntrada(dataAula), normalizar(nome || '')].join('|'),
    porId: [normalizarDataEntrada(dataAula), normalizar(alunoId || '')].join('|'),
    porNomeTurno: [normalizarDataEntrada(dataAula), normalizar(nome || ''), normalizar(turno || '')].join('|'),
    porRgCpf: [normalizarDataEntrada(dataAula), normalizar(rgCpf || '')].join('|')
  };
}


function extrairDataDeDataHora(valor) {
  const txt = valorSeguro(valor);
  if (!txt) return normalizarDataEntrada('');

  const m = txt.match(/^(\d{2}\/\d{2}\/\d{4})/);
  if (m) return m[1];

  return normalizarDataEntrada(txt);
}

function gerarTrIdCurto() {
  return Utilities.getUuid().replace(/-/g, '').substring(0, 8);
}

function gerarHashSenha(senha) {
  const bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    String(senha || ''),
    Utilities.Charset.UTF_8
  );

  return bytes.map(function(b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}


function trocarSenhaPrimeiroAcesso(token, payload) {
  try {
    const sessao = validarSessao(token);
    payload = payload || {};

    const novaSenha = String(payload.novaSenha || '').trim();
    const confirmarSenha = String(payload.confirmarSenha || '').trim();

    if (!novaSenha || novaSenha.length < 4) {
      throw new Error('A nova senha deve ter pelo menos 4 caracteres.');
    }
    if (novaSenha !== confirmarSenha) {
      throw new Error('A confirmação da senha não confere.');
    }

    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName(ABA_OPERADORES);
    if (!aba) throw new Error('A aba OPERADORES não foi encontrada.');

    const dados = aba.getDataRange().getDisplayValues();
    const mapa = mapearCabecalhos(dados[0]);

    const idxSenha = pegarIndiceOpcional(mapa, ['senha']);
    const idxSenhaHash = pegarIndiceOpcional(mapa, ['senhahash']);
    const idxTrocar = pegarIndiceObrigatorio(mapa, ['trocarsenha']);

    if (idxSenha > -1) aba.getRange(sessao.rowOperador, idxSenha + 1).setValue('');
    if (idxSenhaHash > -1) aba.getRange(sessao.rowOperador, idxSenhaHash + 1).setValue(gerarHashSenha(novaSenha));
    aba.getRange(sessao.rowOperador, idxTrocar + 1).setValue('NÃO');

    sessao.trocarSenha = false;
    CacheService.getScriptCache().put(
      'sessao_' + token,
      JSON.stringify(sessao),
      TEMPO_SESSAO_HORAS * 3600
    );

    registrarLogInterno({
      usuario: sessao.usuario,
      nome: sessao.nome,
      perfil: sessao.perfil,
      acao: 'TROCOU_SENHA',
      detalhes: 'Senha alterada no primeiro acesso.'
    });

    return { ok: true, mensagem: 'Senha alterada com sucesso.' };
  } catch (erro) {
    return { ok: false, mensagem: erro.message || 'Falha ao trocar a senha.' };
  }
}

function listarOperadoresPainel(token) {
  try {
    const sessao = validarSessao(token);
    if (sessao.perfil !== 'ADMIN') {
      throw new Error('Você não tem permissão para visualizar operadores.');
    }

    garantirEstruturaOperadores();

    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName(ABA_OPERADORES);
    const dados = aba.getDataRange().getDisplayValues();

    if (dados.length < 2) return { ok: true, registros: [] };

    const mapa = mapearCabecalhos(dados[0]);
    const idxMatricula = pegarIndiceObrigatorio(mapa, ['matricula']);
    const idxNome = pegarIndiceObrigatorio(mapa, ['nome']);
    const idxPerfil = pegarIndiceObrigatorio(mapa, ['perfil']);
    const idxStatus = pegarIndiceObrigatorio(mapa, ['status']);
    const idxTrocar = pegarIndiceObrigatorio(mapa, ['trocarsenha']);

    const registros = [];
    for (let i = 1; i < dados.length; i++) {
      const row = dados[i];
      if (!valorSeguro(row[idxMatricula]) && !valorSeguro(row[idxNome])) continue;

      const idxCodigoServidor = pegarIndiceOpcional(mapa, ['codigoservidor']);
      registros.push({
        matricula: valorSeguro(row[idxMatricula]),
        nome: valorSeguro(row[idxNome]),
        perfil: valorSeguro(row[idxPerfil]),
        status: valorSeguro(row[idxStatus]),
        trocarSenha: valorSeguro(row[idxTrocar]),
        codigoServidor: idxCodigoServidor > -1 ? valorSeguro(row[idxCodigoServidor]) : ''
      });
    }

    return { ok: true, registros: registros };
  } catch (erro) {
    return { ok: false, mensagem: erro.message || 'Falha ao listar operadores.' };
  }
}

function salvarOperadorPainel(token, payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    const sessao = validarSessao(token);
    if (sessao.perfil !== 'ADMIN') {
      throw new Error('Você não tem permissão para cadastrar operadores.');
    }

    garantirEstruturaOperadores();

    payload = payload || {};
    const matricula = String(payload.matricula || '').trim();
    const nome = String(payload.nome || '').trim();
    const senha = String(payload.senha || '').trim();
    const perfil = String(payload.perfil || 'OPERADOR').trim().toUpperCase();
    const status = String(payload.status || 'ATIVO').trim().toUpperCase();
    const trocarSenha = String(payload.trocarSenha || 'SIM').trim().toUpperCase();

    if (!matricula) throw new Error('Informe a matrícula.');
    if (!nome) throw new Error('Informe o nome.');
    if (!senha || senha.length < 4) throw new Error('A senha deve ter pelo menos 4 caracteres.');
    if (['ADMIN', 'OPERADOR'].indexOf(perfil) === -1) throw new Error('Perfil inválido.');
    if (['ATIVO', 'INATIVO'].indexOf(status) === -1) throw new Error('Status inválido.');
    if (['SIM', 'NÃO', 'NAO'].indexOf(trocarSenha) === -1) throw new Error('Valor inválido para TROCAR_SENHA.');

    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName(ABA_OPERADORES);
    const dados = aba.getDataRange().getDisplayValues();
    const mapa = mapearCabecalhos(dados[0]);

    const idxMatricula = pegarIndiceObrigatorio(mapa, ['matricula']);
    const idxNome = pegarIndiceObrigatorio(mapa, ['nome']);
    const idxSenha = pegarIndiceOpcional(mapa, ['senha']);
    const idxPerfil = pegarIndiceObrigatorio(mapa, ['perfil']);
    const idxStatus = pegarIndiceObrigatorio(mapa, ['status']);
    const idxTrocar = pegarIndiceObrigatorio(mapa, ['trocarsenha']);
    const idxUsuario = pegarIndiceOpcional(mapa, ['usuario']);
    const idxSenhaHash = pegarIndiceOpcional(mapa, ['senhahash']);
    const idxCriado = pegarIndiceOpcional(mapa, ['criadoem']);
    const idxObs = pegarIndiceOpcional(mapa, ['observacao']);
    const idxCodigoServidor = pegarIndiceOpcional(mapa, ['codigoservidor']);

    for (let i = 1; i < dados.length; i++) {
      const row = dados[i];
      if (normalizar(row[idxMatricula]) === normalizar(matricula)) {
        throw new Error('Já existe um operador com essa matrícula.');
      }
    }

    const totalCols = Math.max(aba.getLastColumn(), dados[0].length);
    const novaLinha = new Array(totalCols).fill('');

    novaLinha[idxMatricula] = matricula;
    novaLinha[idxNome] = nome;
    if (idxSenha > -1) novaLinha[idxSenha] = senha;
    novaLinha[idxPerfil] = perfil;
    novaLinha[idxStatus] = status;
    novaLinha[idxTrocar] = trocarSenha === 'NAO' ? 'NÃO' : trocarSenha;
    if (idxUsuario > -1) novaLinha[idxUsuario] = matricula;
    if (idxSenhaHash > -1) novaLinha[idxSenhaHash] = gerarHashSenha(senha);
    if (idxCriado > -1) novaLinha[idxCriado] = new Date();
    if (idxObs > -1) novaLinha[idxObs] = 'Cadastro interno do painel.';
    if (idxCodigoServidor > -1) novaLinha[idxCodigoServidor] = '';

    aba.appendRow(novaLinha);

    registrarLogInterno({
      usuario: sessao.usuario,
      nome: sessao.nome,
      perfil: sessao.perfil,
      acao: 'CADASTROU_OPERADOR',
      detalhes: 'Novo operador cadastrado: ' + nome + ' (' + matricula + ').'
    });

    return { ok: true, mensagem: 'Operador cadastrado com sucesso.' };
  } catch (erro) {
    return { ok: false, mensagem: erro.message || 'Falha ao cadastrar operador.' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


function atualizarOperadorPainel(token, payload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    const sessao = validarSessao(token);
    if (sessao.perfil !== 'ADMIN') {
      throw new Error('Você não tem permissão para editar operadores.');
    }

    garantirEstruturaOperadores();

    payload = payload || {};
    const matriculaOriginal = String(payload.matriculaOriginal || '').trim();
    const matricula = String(payload.matricula || '').trim();
    const nome = String(payload.nome || '').trim();
    const senha = String(payload.senha || '').trim();
    const perfil = String(payload.perfil || 'OPERADOR').trim().toUpperCase();
    const status = String(payload.status || 'ATIVO').trim().toUpperCase();
    const trocarSenha = String(payload.trocarSenha || 'SIM').trim().toUpperCase();
    const codigoServidor = String(payload.codigoServidor || '').trim();

    if (!matriculaOriginal) throw new Error('Matrícula original não informada.');
    if (!matricula) throw new Error('Informe a matrícula.');
    if (!nome) throw new Error('Informe o nome.');
    if (['ADMIN', 'OPERADOR'].indexOf(perfil) === -1) throw new Error('Perfil inválido.');
    if (['ATIVO', 'INATIVO'].indexOf(status) === -1) throw new Error('Status inválido.');
    if (['SIM', 'NÃO', 'NAO'].indexOf(trocarSenha) === -1) throw new Error('Valor inválido para TROCAR_SENHA.');

    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const aba = ss.getSheetByName(ABA_OPERADORES);
    const dados = aba.getDataRange().getDisplayValues();
    const mapa = mapearCabecalhos(dados[0]);

    const idxMatricula = pegarIndiceObrigatorio(mapa, ['matricula']);
    const idxNome = pegarIndiceObrigatorio(mapa, ['nome']);
    const idxSenha = pegarIndiceOpcional(mapa, ['senha']);
    const idxPerfil = pegarIndiceObrigatorio(mapa, ['perfil']);
    const idxStatus = pegarIndiceObrigatorio(mapa, ['status']);
    const idxTrocar = pegarIndiceObrigatorio(mapa, ['trocarsenha']);
    const idxUsuario = pegarIndiceOpcional(mapa, ['usuario']);
    const idxSenhaHash = pegarIndiceOpcional(mapa, ['senhahash']);
    const idxObs = pegarIndiceOpcional(mapa, ['observacao']);
    const idxCodigoServidor = pegarIndiceOpcional(mapa, ['codigoservidor']);

    let linhaAlvo = -1;

    for (let i = 1; i < dados.length; i++) {
      const row = dados[i];
      const mat = valorSeguro(row[idxMatricula]);

      if (normalizar(mat) === normalizar(matriculaOriginal)) {
        linhaAlvo = i + 1;
      } else if (normalizar(mat) === normalizar(matricula)) {
        throw new Error('Já existe outro operador com essa matrícula.');
      }
    }

    if (linhaAlvo === -1) throw new Error('Operador não encontrado.');

    aba.getRange(linhaAlvo, idxMatricula + 1).setValue(matricula);
    aba.getRange(linhaAlvo, idxNome + 1).setValue(nome);
    aba.getRange(linhaAlvo, idxPerfil + 1).setValue(perfil);
    aba.getRange(linhaAlvo, idxStatus + 1).setValue(status);
    aba.getRange(linhaAlvo, idxTrocar + 1).setValue(trocarSenha === 'NAO' ? 'NÃO' : trocarSenha);

    if (idxUsuario > -1) aba.getRange(linhaAlvo, idxUsuario + 1).setValue(matricula);
    if (idxCodigoServidor > -1) aba.getRange(linhaAlvo, idxCodigoServidor + 1).setValue(codigoServidor);
    if (idxObs > -1) aba.getRange(linhaAlvo, idxObs + 1).setValue('Operador atualizado pelo painel.');

    if (senha) {
      if (senha.length < 4) throw new Error('A nova senha deve ter pelo menos 4 caracteres.');
      if (idxSenha > -1) aba.getRange(linhaAlvo, idxSenha + 1).setValue(senha);
      if (idxSenhaHash > -1) aba.getRange(linhaAlvo, idxSenhaHash + 1).setValue(gerarHashSenha(senha));
    }

    registrarLogInterno({
      usuario: sessao.usuario,
      nome: sessao.nome,
      perfil: sessao.perfil,
      acao: 'ATUALIZOU_OPERADOR',
      detalhes: 'Operador atualizado: ' + nome + ' (' + matriculaOriginal + ' → ' + matricula + ').'
    });

    return { ok: true, mensagem: 'Operador atualizado com sucesso.' };
  } catch (erro) {
    return { ok: false, mensagem: erro.message || 'Falha ao atualizar operador.' };
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}


function listarHistoricoAlunoPainel(token, payload) {
  try {
    validarSessao(token);
    payload = payload || {};

    const alunoId = valorSeguro(payload.alunoId);
    const nome = valorSeguro(payload.nome);
    const rgCpf = valorSeguro(payload.rgCpf);
    const turno = valorSeguro(payload.turno);
    const horario = valorSeguro(payload.horario);
    const dias = valorSeguro(payload.dias);
    const modalidade = valorSeguro(payload.modalidade);

    const ss = SpreadsheetApp.openById(ID_PLANILHA);
    const abaFreq = ss.getSheetByName(ABA_FREQUENCIA);
    if (!abaFreq || abaFreq.getLastRow() < 2) {
      return { ok: true, registros: [] };
    }

    const dados = abaFreq.getDataRange().getDisplayValues();
    const cab = dados[0];
    const mapa = mapearCabecalhos(cab);

    const idxDataHora   = pegarIndiceOpcional(mapa, ['datahora']);
    const idxId         = pegarIndiceOpcional(mapa, ['trid', 'id']);
    const idxNome       = pegarIndiceOpcional(mapa, ['nome', 'nomedoaluno', 'aluno']);
    const idxRGCPF      = pegarIndiceOpcional(mapa, ['rgcpf', 'rg', 'cpf']);
    const idxModalidade = pegarIndiceOpcional(mapa, ['modalidade']);
    const idxTurno      = pegarIndiceOpcional(mapa, ['turno']);
    const idxHorario    = pegarIndiceOpcional(mapa, ['horario']);
    const idxDias       = pegarIndiceOpcional(mapa, ['dias']);
    const idxStatus     = pegarIndiceOpcional(mapa, ['presenteausente']);
    const idxJustific   = pegarIndiceOpcional(mapa, ['justificarausencia']);
    const idxPcd        = pegarIndiceOpcional(mapa, ['epcd', 'pcd']);
    const idxCid        = pegarIndiceOpcional(mapa, ['sesiminformeocid', 'cid']);
    const idxRespons    = pegarIndiceOpcional(mapa, ['nomedoresponsavel', 'responsavel']);
    const idxCel        = pegarIndiceOpcional(mapa, ['celwhatsapp', 'celular', 'whatsapp']);

    const registros = [];

    for (let i = 1; i < dados.length; i++) {
      const row = dados[i];

      const rowObj = {
        alunoIdLinha: idxId > -1 ? valorSeguro(row[idxId]) : '',
        nomeLinha: idxNome > -1 ? valorSeguro(row[idxNome]) : '',
        rgCpfLinha: idxRGCPF > -1 ? valorSeguro(row[idxRGCPF]) : '',
        modalidadeLinha: idxModalidade > -1 ? valorSeguro(row[idxModalidade]) : '',
        turnoLinha: idxTurno > -1 ? valorSeguro(row[idxTurno]) : '',
        horarioLinha: idxHorario > -1 ? valorSeguro(row[idxHorario]) : '',
        diasLinha: idxDias > -1 ? valorSeguro(row[idxDias]) : ''
      };

      const mesmoAluno =
        (alunoId && rowObj.alunoIdLinha && normalizar(alunoId) === normalizar(rowObj.alunoIdLinha)) ||
        (rgCpf && rowObj.rgCpfLinha && normalizar(rgCpf) === normalizar(rowObj.rgCpfLinha)) ||
        (nome && rowObj.nomeLinha && normalizar(nome) === normalizar(rowObj.nomeLinha));

      if (!mesmoAluno) continue;

      if (modalidade && rowObj.modalidadeLinha && normalizar(modalidade) !== normalizar(rowObj.modalidadeLinha)) continue;
      if (turno && rowObj.turnoLinha && normalizar(turno) !== normalizar(rowObj.turnoLinha)) continue;
      if (horario && rowObj.horarioLinha && normalizar(horario) !== normalizar(rowObj.horarioLinha)) continue;
      if (dias && rowObj.diasLinha && normalizar(dias) !== normalizar(rowObj.diasLinha)) continue;

      registros.push({
        dataHora: idxDataHora > -1 ? valorSeguro(row[idxDataHora]) : '',
        status: idxStatus > -1 ? valorSeguro(row[idxStatus]) : '',
        justificativa: idxJustific > -1 ? valorSeguro(row[idxJustific]) : '',
        modalidade: rowObj.modalidadeLinha,
        turno: rowObj.turnoLinha,
        horario: rowObj.horarioLinha,
        dias: rowObj.diasLinha,
        pcd: idxPcd > -1 ? valorSeguro(row[idxPcd]) : '',
        cid: idxCid > -1 ? valorSeguro(row[idxCid]) : '',
        responsavel: idxRespons > -1 ? valorSeguro(row[idxRespons]) : '',
        celular: idxCel > -1 ? valorSeguro(row[idxCel]) : ''
      });
    }

    registros.sort(function(a, b) {
      return obterDataHoraOrdenavel(b.dataHora) - obterDataHoraOrdenavel(a.dataHora);
    });

    return { ok: true, registros: registros };
  } catch (erro) {
    return { ok: false, mensagem: erro.message || 'Falha ao carregar histórico do aluno.' };
  }
}


function carregarPainelInicial(token) {
  try {
    const filtrosRes = obterFiltrosPainel(token);
    if (!filtrosRes || !filtrosRes.ok) return filtrosRes;

    const consultaRes = buscarPainel(token, {
      dataAula: normalizarDataEntrada(new Date()),
      pagina: 1,
      limite: LIMITE_CONSULTA_PADRAO
    });
    if (!consultaRes || !consultaRes.ok) return consultaRes;

    return {
      ok: true,
      filtros: filtrosRes.filtros,
      consulta: consultaRes
    };
  } catch (erro) {
    return { ok: false, mensagem: erro.message || 'Falha ao inicializar o painel.' };
  }
}

const APP = {
  CONFIG: 'CONFIG',

  FUNC_START_ROW: 5,
  FERIAS_START_ROW: 5,

  CORES: {
    OK_BG: '#d9ead3',
    OK_TEXT: '#38761d',

    OCORRENCIA_BG: '#f4cccc',
    OCORRENCIA_TEXT: '#990000',

    MEDICO_BG: '#fff2cc',
    MEDICO_TEXT: '#7f6000',

    FERIAS_BG: '#e2f0d9',
    FERIAS_TEXT: '#38761d',

    FIM_SEMANA_BG: '#eeeeee',
    FIM_SEMANA_TEXT: '#777777',

    CABECALHO: '#1e3a8a',
    CABECALHO_2: '#2563eb',

    FUNCIONARIOS: [
      '#dbeafe',
      '#cfe2f3',
      '#c5d9f1',
      '#b9cde5',
      '#aec4df',
      '#a4bdd8'
    ]
  }
};


/************************************************************
 * MENU
 ************************************************************/

function onOpen() {

  SpreadsheetApp.getUi()
    .createMenu('📋 Controle de Ponto')
    .addItem('🎛️ Abrir painel', 'abrirPainel')
    .addItem('⚙️ Configurar sistema', 'setupPlanilha')
    .addItem('📅 Gerar / atualizar mês', 'abrirPainel')
    .addItem('📊 Abrir relatório', 'abrirAbaRelatorio')
    .addToUi();
}


/************************************************************
 * ABRIR PAINEL
 ************************************************************/

function abrirPainel() {

  const html = HtmlService
    .createHtmlOutputFromFile('Painel')
    .setTitle('Controle de Ponto');

  SpreadsheetApp.getUi()
    .showSidebar(html);
}


/************************************************************
 * CONFIGURAÇÃO INICIAL
 ************************************************************/

function setupPlanilha() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let config = ss.getSheetByName(APP.CONFIG);

  if (!config) {
    config = ss.insertSheet(APP.CONFIG);
  }

  configurarCabecalhos_(config);

  const primeiraCelula =
    String(config.getRange('A5').getValue()).trim();

  if (!primeiraCelula) {

    const funcionarios = [
      ['Dani', '08:00', '12:00', '13:00', '17:00', 'Ativo'],
      ['Karina', '08:00', '12:00', '13:00', '17:00', 'Ativo'],
      ['Matheus', '08:00', '12:00', '13:00', '17:00', 'Ativo'],
      ['Thamirys', '08:00', '12:00', '13:00', '17:00', 'Ativo'],

      ['Ivanete', '07:00', '11:00', '12:00', '16:00', 'Ativo'],
      ['Bel', '07:00', '11:00', '12:00', '16:00', 'Ativo'],
      ['Cida', '07:00', '11:00', '12:00', '16:00', 'Ativo'],
      ['Aronita', '07:00', '11:00', '12:00', '16:00', 'Ativo'],

      ['Laís', '11:00', '17:00', '', '', 'Ativo'],
      ['Vanessa', '07:00', '11:00', '', '', 'Ativo']
    ];

    config
      .getRange(5, 1, funcionarios.length, 6)
      .setValues(funcionarios);
  }

  formatarConfig_(config);

  criarRelatorio_();

  SpreadsheetApp.getUi().alert(
    'Sistema configurado com sucesso.\n\n' +
    'Funcionários, horários e relatório estão disponíveis no painel.'
  );
}


/************************************************************
 * CABEÇALHOS DA CONFIG
 ************************************************************/

function configurarCabecalhos_(sheet) {

  sheet.getRange('A1').setValue(
    'CONFIGURAÇÃO DO CONTROLE DE PONTO'
  );

  sheet.getRange('A3:F3').setValues([[
    'Funcionário',
    'Entrada',
    'Saída 1',
    'Entrada 2',
    'Saída 2',
    'Situação'
  ]]);

  sheet.getRange('H3:K3').setValues([[
    'Funcionário',
    'Início',
    'Fim',
    'Observação'
  ]]);
}


/************************************************************
 * FORMATAÇÃO DA CONFIG
 ************************************************************/

function formatarConfig_(sheet) {

  const titulo = sheet.getRange('A1:K1');

  titulo.breakApart();

  titulo
    .merge()
    .setBackground(APP.CORES.CABECALHO)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(14)
    .setHorizontalAlignment('center');

  sheet.getRange('A3:F3')
    .setBackground(APP.CORES.CABECALHO_2)
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  sheet.getRange('H3:K3')
    .setBackground(APP.CORES.CABECALHO_2)
    .setFontColor('#ffffff')
    .setFontWeight('bold');

  sheet.setFrozenRows(3);

  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 90);
  sheet.setColumnWidth(3, 90);
  sheet.setColumnWidth(4, 90);
  sheet.setColumnWidth(5, 90);
  sheet.setColumnWidth(6, 90);

  sheet.setColumnWidth(8, 140);
  sheet.setColumnWidth(9, 100);
  sheet.setColumnWidth(10, 100);
  sheet.setColumnWidth(11, 200);
}


/************************************************************
 * OBTER CONFIG
 ************************************************************/

function getConfig_() {

  const ss = SpreadsheetApp.getActiveSpreadsheet();

  let sheet = ss.getSheetByName(APP.CONFIG);

  if (!sheet) {

    setupPlanilha();

    sheet = ss.getSheetByName(APP.CONFIG);
  }

  if (!sheet) {
    throw new Error('A aba CONFIG não foi encontrada.');
  }

  return sheet;
}


/************************************************************
 * FUNCIONÁRIOS
 ************************************************************/

function getFuncionarios() {

  const sheet = getConfig_();

  const ultimaLinha =
    sheet.getLastRow();

  if (ultimaLinha < APP.FUNC_START_ROW) {
    return [];
  }

  const quantidade =
    ultimaLinha - APP.FUNC_START_ROW + 1;

  const valores =
    sheet
      .getRange(
        APP.FUNC_START_ROW,
        1,
        quantidade,
        6
      )
      .getDisplayValues();

  const funcionarios = [];

  valores.forEach(function(linha) {

    const nome =
      String(linha[0] || '').trim();

    if (!nome) {
      return;
    }

    funcionarios.push({
      nome: nome,
      entrada: normalizarHorario_(linha[1]),
      saida1: normalizarHorario_(linha[2]),
      entrada2: normalizarHorario_(linha[3]),
      saida2: normalizarHorario_(linha[4]),
      situacao: String(linha[5] || 'Ativo').trim()
    });
  });

  return funcionarios;
}


/************************************************************
 * NORMALIZAR HORÁRIO
 ************************************************************/

function normalizarHorario_(valor) {

  if (valor === null || valor === undefined) {
    return '';
  }

  let texto = String(valor).trim();

  if (!texto) {
    return '';
  }

  if (/^\d{1,2}:\d{2}:\d{2}$/.test(texto)) {
    return texto.substring(0, 5);
  }

  if (/^\d{1,2}:\d{2}$/.test(texto)) {

    const partes = texto.split(':');

    return (
      String(partes[0]).padStart(2, '0') +
      ':' +
      partes[1]
    );
  }

  return texto;
}


/************************************************************
 * DADOS INICIAIS DO PAINEL
 ************************************************************/

function getDadosIniciais() {

  const agora = new Date();

  return {
    mesAtual: agora.getMonth() + 1,
    anoAtual: agora.getFullYear(),
    funcionarios: getFuncionarios()
  };
}


/************************************************************
 * HORÁRIOS DE UM FUNCIONÁRIO
 ************************************************************/

function getHorariosFuncionario_(nome) {

  const funcionarios =
    getFuncionarios();

  const funcionario =
    funcionarios.find(function(f) {
      return f.nome === nome;
    });

  if (!funcionario) {
    return [];
  }

  const horarios = [];

  if (funcionario.entrada) {
    horarios.push({
      nome: 'Entrada',
      horario: funcionario.entrada
    });
  }

  if (funcionario.saida1) {
    horarios.push({
      nome: 'Saída 1',
      horario: funcionario.saida1
    });
  }

  if (funcionario.entrada2) {
    horarios.push({
      nome: 'Entrada 2',
      horario: funcionario.entrada2
    });
  }

  if (funcionario.saida2) {
    horarios.push({
      nome: 'Saída 2',
      horario: funcionario.saida2
    });
  }

  return horarios;
}


/************************************************************
 * SALVAR HORÁRIO
 ************************************************************/

function salvarHorario(dados) {

  if (!dados || !dados.funcionario) {
    throw new Error('Funcionário não informado.');
  }

  const sheet = getConfig_();

  const ultimaLinha =
    Math.max(
      sheet.getLastRow(),
      APP.FUNC_START_ROW
    );

  const valores =
    sheet
      .getRange(
        APP.FUNC_START_ROW,
        1,
        ultimaLinha - APP.FUNC_START_ROW + 1,
        6
      )
      .getDisplayValues();

  let linhaEncontrada = -1;

  for (let i = 0; i < valores.length; i++) {

    if (
      String(valores[i][0]).trim() ===
      String(dados.funcionario).trim()
    ) {

      linhaEncontrada =
        APP.FUNC_START_ROW + i;

      break;
    }
  }

  if (linhaEncontrada === -1) {
    throw new Error(
      'Funcionário não encontrado na CONFIG: ' +
      dados.funcionario
    );
  }

  sheet
    .getRange(linhaEncontrada, 2, 1, 4)
    .setValues([[
      dados.entrada || '',
      dados.saida1 || '',
      dados.entrada2 || '',
      dados.saida2 || ''
    ]]);

  return 'Horário atualizado com sucesso.';
}


/************************************************************
 * FÉRIAS
 ************************************************************/

function salvarFerias(dados) {

  if (!dados.funcionario) {
    throw new Error('Funcionário não informado.');
  }

  if (!dados.inicio || !dados.fim) {
    throw new Error('Informe início e fim das férias.');
  }

  const sheet = getConfig_();

  let linha = sheet.getLastRow() + 1;

  if (linha < APP.FERIAS_START_ROW) {
    linha = APP.FERIAS_START_ROW;
  }

  sheet
    .getRange(linha, 8, 1, 4)
    .setValues([[
      dados.funcionario,
      converterData_(dados.inicio),
      converterData_(dados.fim),
      dados.observacao || ''
    ]]);

  sheet
    .getRange(linha, 9, 1, 2)
    .setNumberFormat('dd/MM/yyyy');

  return 'Férias registradas com sucesso.';
}


/************************************************************
 * LER FÉRIAS
 ************************************************************/

function getFerias_() {

  const sheet = getConfig_();

  const ultimaLinha =
    sheet.getLastRow();

  if (ultimaLinha < APP.FERIAS_START_ROW) {
    return [];
  }

  const quantidade =
    ultimaLinha - APP.FERIAS_START_ROW + 1;

  const valores =
    sheet
      .getRange(
        APP.FERIAS_START_ROW,
        8,
        quantidade,
        4
      )
      .getValues();

  return valores
    .filter(function(linha) {
      return String(linha[0] || '').trim() !== '';
    })
    .map(function(linha) {

      return {
        funcionario: String(linha[0]).trim(),
        inicio: linha[1],
        fim: linha[2],
        observacao: linha[3] || ''
      };

    });
}


/************************************************************
 * SALVAR OCORRÊNCIA
 *
 * ATESTADO:
 * Cria uma ocorrência para TODOS os horários
 * daquele funcionário naquele dia.
 ************************************************************/

function salvarOcorrencia(dados) {

  if (!dados) {
    throw new Error('Dados da ocorrência não recebidos.');
  }

  if (!dados.funcionario) {
    throw new Error('Selecione um funcionário.');
  }

  if (!dados.data) {
    throw new Error('Informe a data.');
  }

  /*
   * Para Atestado não exigimos registro,
   * pois ele vale para o dia inteiro.
   */
  if (
    dados.situacao !== 'Atestado' &&
    !dados.registro
  ) {
    throw new Error('Selecione o registro.');
  }

  if (!dados.motivo) {
    throw new Error('Informe o motivo.');
  }

  const data =
    converterData_(dados.data);

  const mes =
    data.getMonth() + 1;

  const ano =
    data.getFullYear();

  const sheet =
    obterOuCriarMes_(mes, ano);

  let ocorrencias =
    lerOcorrenciasAba_(sheet);

  const funcionario =
    getFuncionarios().find(function(f) {
      return f.nome === dados.funcionario;
    });

  if (!funcionario) {
    throw new Error(
      'Funcionário não encontrado: ' +
      dados.funcionario
    );
  }

  const horarios =
    getHorariosFuncionario_(dados.funcionario);


  /**********************************************************
   * ATESTADO = DIA INTEIRO
   **********************************************************/

  if (dados.situacao === 'Atestado') {

    if (!horarios.length) {
      throw new Error(
        'O funcionário não possui horários cadastrados.'
      );
    }

    /*
     * Remove todas as ocorrências anteriores
     * daquele funcionário naquele dia.
     */
    ocorrencias =
      ocorrencias.filter(function(o) {

        return !(
          o.funcionario === dados.funcionario &&
          mesmaData_(o.data, data)
        );

      });


    /*
     * Cria uma ocorrência para cada horário.
     *
     * Isso é importante para que a aba mensal
     * mostre ATESTADO em todos os horários.
     */
    horarios.forEach(function(horario) {

      ocorrencias.push({

        data:
          dados.data,

        funcionario:
          dados.funcionario,

        registro:
          horario.nome,

        horarioPrevisto:
          horario.horario,

        horarioInformado:
          '',

        situacao:
          'Atestado',

        motivo:
          dados.motivo,

        observacao:
          dados.observacao || ''

      });

    });

  } else {

    /*
     * OCORRÊNCIA NORMAL
     */

    const horarioSelecionado =
      horarios.find(function(h) {
        return h.nome === dados.registro;
      });

    const horarioPrevisto =
      horarioSelecionado
        ? horarioSelecionado.horario
        : '';

    const novaOcorrencia = {

      data:
        dados.data,

      funcionario:
        dados.funcionario,

      registro:
        dados.registro,

      horarioPrevisto:
        horarioPrevisto,

      horarioInformado:
        dados.horario || '',

      situacao:
        dados.situacao || 'Outro',

      motivo:
        dados.motivo || '',

      observacao:
        dados.observacao || ''
    };

    const indice =
      ocorrencias.findIndex(function(o) {

        return (
          mesmaData_(o.data, data) &&
          o.funcionario === novaOcorrencia.funcionario &&
          o.registro === novaOcorrencia.registro
        );

      });

    if (indice >= 0) {

      ocorrencias[indice] =
        novaOcorrencia;

    } else {

      ocorrencias.push(
        novaOcorrencia
      );

    }

  }

  renderizarMes_(
    sheet,
    mes,
    ano,
    ocorrencias
  );

  atualizarRelatorio_();

  if (dados.situacao === 'Atestado') {
    return 'Atestado registrado para o dia inteiro.';
  }

  return 'Ocorrência salva com sucesso.';
}


/************************************************************
 * CONVERTER DATA
 ************************************************************/

function converterData_(valor) {

  if (valor instanceof Date) {
    return new Date(valor);
  }

  const texto =
    String(valor).trim();

  if (/^\d{4}-\d{2}-\d{2}$/.test(texto)) {

    const partes =
      texto.split('-');

    return new Date(
      Number(partes[0]),
      Number(partes[1]) - 1,
      Number(partes[2])
    );
  }

  if (/^\d{2}\/\d{2}\/\d{4}$/.test(texto)) {

    const partes =
      texto.split('/');

    return new Date(
      Number(partes[2]),
      Number(partes[1]) - 1,
      Number(partes[0])
    );
  }

  throw new Error(
    'Data inválida: ' + texto
  );
}


/************************************************************
 * OBTER OU CRIAR MÊS
 ************************************************************/

function obterOuCriarMes_(mes, ano) {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  const nome =
    nomeMes_(mes) + '-' + ano;

  let sheet =
    ss.getSheetByName(nome);

  if (!sheet) {
    sheet =
      ss.insertSheet(nome);
  }

  return sheet;
}


/************************************************************
 * GERAR MÊS
 ************************************************************/

function gerarMes(mes, ano) {

  mes = Number(mes);
  ano = Number(ano);

  if (
    !mes ||
    mes < 1 ||
    mes > 12
  ) {
    throw new Error('Mês inválido.');
  }

  if (
    !ano ||
    ano < 2000 ||
    ano > 2100
  ) {
    throw new Error('Ano inválido.');
  }

  const sheet =
    obterOuCriarMes_(mes, ano);

  const ocorrencias =
    lerOcorrenciasAba_(sheet);

  renderizarMes_(
    sheet,
    mes,
    ano,
    ocorrencias
  );

  atualizarRelatorio_();

  return (
    'Mês ' +
    nomeMes_(mes) +
    '/' +
    ano +
    ' gerado com sucesso.'
  );
}


/************************************************************
 * NOME DO MÊS
 ************************************************************/

function nomeMes_(mes) {

  const nomes = [
    'JAN',
    'FEV',
    'MAR',
    'ABR',
    'MAI',
    'JUN',
    'JUL',
    'AGO',
    'SET',
    'OUT',
    'NOV',
    'DEZ'
  ];

  return nomes[Number(mes) - 1];
}


/************************************************************
 * RENDERIZAR MÊS
 ************************************************************/

function renderizarMes_(
  sheet,
  mes,
  ano,
  ocorrencias
) {

  const funcionarios =
    getFuncionarios()
      .filter(function(f) {
        return f.situacao.toLowerCase() !== 'inativo';
      });

  const ferias =
    getFerias_();

  const ocorrenciasSalvas =
    ocorrencias || [];

  sheet.clear();

  const dias =
    new Date(
      ano,
      mes,
      0
    ).getDate();

  const ultimaColuna =
    dias + 1;


  /**********************************************************
   * TÍTULO
   **********************************************************/

  const titulo =
    sheet.getRange(
      1,
      1,
      1,
      ultimaColuna
    );

  titulo
    .breakApart()
    .merge()
    .setValue(
      'CONTROLE DE PONTO - ' +
      nomeMes_(mes) +
      '/' +
      ano
    )
    .setBackground(APP.CORES.CABECALHO)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  sheet.setRowHeight(1, 32);


  /**********************************************************
   * LINHA DAS DATAS
   **********************************************************/

  sheet
    .getRange(4, 1)
    .setValue('Funcionário');

  for (let dia = 1; dia <= dias; dia++) {

    const data =
      new Date(
        ano,
        mes - 1,
        dia
      );

    sheet
      .getRange(4, dia + 1)
      .setValue(
        Utilities.formatDate(
          data,
          Session.getScriptTimeZone(),
          'dd/MM'
        )
      );
  }


  /**********************************************************
   * DIAS DA SEMANA
   **********************************************************/

  sheet
    .getRange(5, 1)
    .setValue('');

  const diasSemana = [
    'DOM',
    'SEG',
    'TER',
    'QUA',
    'QUI',
    'SEX',
    'SÁB'
  ];

  for (let dia = 1; dia <= dias; dia++) {

    const data =
      new Date(
        ano,
        mes - 1,
        dia
      );

    sheet
      .getRange(5, dia + 1)
      .setValue(
        diasSemana[data.getDay()]
      );
  }

  sheet
    .getRange(
      4,
      1,
      2,
      ultimaColuna
    )
    .setBackground(APP.CORES.CABECALHO_2)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');


  /**********************************************************
   * FUNCIONÁRIOS
   **********************************************************/

  let linhaAtual = 6;

  funcionarios.forEach(function(funcionario, indice) {

    const horarios =
      getHorariosFuncionario_(funcionario.nome);

    const quantidadeLinhas =
      Math.max(
        horarios.length,
        1
      );

    const corFuncionario =
      APP.CORES.FUNCIONARIOS[
        indice %
        APP.CORES.FUNCIONARIOS.length
      ];


    /*
     * Área do nome do funcionário.
     */
    const celulaFuncionario =
      sheet.getRange(
        linhaAtual,
        1,
        quantidadeLinhas,
        1
      );

    celulaFuncionario
      .merge()
      .setValue(funcionario.nome)
      .setFontWeight('bold')
      .setVerticalAlignment('middle')
      .setHorizontalAlignment('center')
      .setBackground(corFuncionario);


    /********************************************************
     * CADA HORÁRIO
     ********************************************************/

    horarios.forEach(function(horario, indiceHorario) {

      const linha =
        linhaAtual + indiceHorario;

      sheet
        .getRange(linha, 1)
        .setBackground(corFuncionario);


      /******************************************************
       * DIAS
       ******************************************************/

      for (let dia = 1; dia <= dias; dia++) {

        const data =
          new Date(
            ano,
            mes - 1,
            dia
          );

        const coluna =
          dia + 1;

        const diaSemana =
          data.getDay();


        /*
         * FINAL DE SEMANA
         */
        if (
          diaSemana === 0 ||
          diaSemana === 6
        ) {

          sheet
            .getRange(
              linha,
              coluna
            )
            .setValue('—')
            .setBackground(
              APP.CORES.FIM_SEMANA_BG
            )
            .setFontColor(
              APP.CORES.FIM_SEMANA_TEXT
            )
            .setHorizontalAlignment('center');

          continue;
        }


        /*
         * FÉRIAS
         */
        if (
          estaDeFerias_(
            funcionario.nome,
            data,
            ferias
          )
        ) {

          sheet
            .getRange(
              linha,
              coluna
            )
            .setValue('🏖️ FÉRIAS')
            .setBackground(
              APP.CORES.FERIAS_BG
            )
            .setFontColor(
              APP.CORES.FERIAS_TEXT
            )
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle');

          continue;
        }


        /*
         * OCORRÊNCIA
         */
        const ocorrencia =
          ocorrenciasSalvas.find(function(o) {

            return (
              o.funcionario === funcionario.nome &&
              mesmaData_(o.data, data) &&
              o.registro === horario.nome
            );

          });


        if (ocorrencia) {

          const classe =
            classificarOcorrencia_(
              ocorrencia
            );

          let texto;

          if (
            ocorrencia.situacao ===
            'Atestado'
          ) {

            texto =
              '🟡 ATESTADO';

          } else if (
            classe.tipo === 'medico'
          ) {

            texto =
              '🟡 MÉDICO\n' +
              ocorrencia.motivo;

          } else {

            texto =
              '🔴 OCORRÊNCIA\n' +
              ocorrencia.motivo;
          }


          sheet
            .getRange(
              linha,
              coluna
            )
            .setValue(texto)
            .setBackground(classe.bg)
            .setFontColor(classe.text)
            .setWrap(true)
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle');

        } else {

          /*
           * DIA NORMAL
           */
          sheet
            .getRange(
              linha,
              coluna
            )
            .setValue('OK')
            .setBackground(
              APP.CORES.OK_BG
            )
            .setFontColor(
              APP.CORES.OK_TEXT
            )
            .setHorizontalAlignment('center')
            .setVerticalAlignment('middle');
        }
      }
    });


    linhaAtual +=
      quantidadeLinhas + 1;
  });


  /**********************************************************
   * LARGURAS
   **********************************************************/

  sheet.setColumnWidth(1, 130);

  for (
    let coluna = 2;
    coluna <= ultimaColuna;
    coluna++
  ) {

    sheet.setColumnWidth(
      coluna,
      90
    );
  }


  /**********************************************************
   * DETALHES DAS OCORRÊNCIAS
   **********************************************************/

  const linhaDetalhes =
    linhaAtual + 1;

  sheet
    .getRange(
      linhaDetalhes,
      1,
      1,
      8
    )
    .breakApart()
    .merge()
    .setValue(
      '📌 DETALHES DAS OCORRÊNCIAS'
    )
    .setBackground(
      APP.CORES.CABECALHO
    )
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(13)
    .setHorizontalAlignment('center');


  sheet
    .getRange(
      linhaDetalhes + 1,
      1,
      1,
      8
    )
    .setValues([[
      'Data',
      'Funcionário',
      'Registro',
      'Horário previsto',
      'Horário informado',
      'Situação',
      'Motivo',
      'Observação'
    ]])
    .setBackground(
      APP.CORES.CABECALHO_2
    )
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center');


  if (ocorrenciasSalvas.length) {

    const dados =
      ocorrenciasSalvas.map(function(o) {

        return [
          o.data,
          o.funcionario,
          o.registro,
          o.horarioPrevisto,
          o.horarioInformado,
          o.situacao,
          o.motivo,
          o.observacao
        ];

      });

    sheet
      .getRange(
        linhaDetalhes + 2,
        1,
        dados.length,
        8
      )
      .setValues(dados);

    sheet
      .getRange(
        linhaDetalhes + 2,
        1,
        dados.length,
        8
      )
      .setVerticalAlignment('middle')
      .setWrap(true);
  }


  /**********************************************************
   * LINKS DAS OCORRÊNCIAS
   **********************************************************/

  aplicarLinksOcorrencias_(
    sheet,
    funcionarios,
    horariosPorFuncionario_(funcionarios),
    ocorrenciasSalvas,
    linhaDetalhes + 2,
    dias,
    ano,
    mes
  );


  /**********************************************************
   * CONFIGURAÇÕES FINAIS
   **********************************************************/

  sheet.setFrozenRows(5);

  sheet
    .getRange(
      4,
      1,
      linhaDetalhes + 1,
      ultimaColuna
    )
    .setVerticalAlignment('middle');

  sheet
    .getRange(
      linhaDetalhes + 1,
      1,
      1,
      8
    )
    .setWrap(true);
}


/************************************************************
 * HORÁRIOS DOS FUNCIONÁRIOS
 ************************************************************/

function horariosPorFuncionario_(funcionarios) {

  const resultado = {};

  funcionarios.forEach(function(f) {

    resultado[f.nome] =
      getHorariosFuncionario_(f.nome);

  });

  return resultado;
}


/************************************************************
 * APLICAR LINKS AOS DETALHES
 ************************************************************/

function aplicarLinksOcorrencias_(
  sheet,
  funcionarios,
  horariosMap,
  ocorrencias,
  linhaInicialDetalhes,
  dias,
  ano,
  mes
) {

  if (!ocorrencias.length) {
    return;
  }

  const baseUrl =
    SpreadsheetApp
      .getActiveSpreadsheet()
      .getUrl() +
    '#gid=' +
    sheet.getSheetId();

  funcionarios.forEach(function(funcionario) {

    const horarios =
      horariosMap[funcionario.nome] || [];

    horarios.forEach(function(horario, indiceHorario) {

      const linhaFuncionario =
        encontrarLinhaFuncionario_(
          funcionarios,
          funcionario.nome,
          horariosMap,
          indiceHorario
        );

      for (let dia = 1; dia <= dias; dia++) {

        const data =
          new Date(
            ano,
            mes - 1,
            dia
          );

        const ocorrencia =
          ocorrencias.find(function(o) {

            return (
              o.funcionario === funcionario.nome &&
              mesmaData_(o.data, data) &&
              o.registro === horario.nome
            );

          });

        if (!ocorrencia) {
          continue;
        }

        const linhaDetalhe =
          encontrarLinhaDetalhePorDados_(
            ocorrencias,
            ocorrencia,
            linhaInicialDetalhes
          );

        if (!linhaDetalhe) {
          continue;
        }

        const coluna =
          dia + 1;

        const url =
          baseUrl +
          '&range=A' +
          linhaDetalhe;

        const celula =
          sheet.getRange(
            linhaFuncionario,
            coluna
          );

        const texto =
          celula.getDisplayValue();

        const rich =
          SpreadsheetApp
            .newRichTextValue()
            .setText(texto)
            .setLinkUrl(url)
            .build();

        celula.setRichTextValue(rich);
      }
    });
  });
}


/************************************************************
 * ENCONTRAR LINHA DO FUNCIONÁRIO
 ************************************************************/

function encontrarLinhaFuncionario_(
  funcionarios,
  nome,
  horariosMap,
  indiceHorario
) {

  let linha = 6;

  for (let i = 0; i < funcionarios.length; i++) {

    const funcionario =
      funcionarios[i];

    const horarios =
      horariosMap[funcionario.nome] || [];

    if (funcionario.nome === nome) {
      return linha + indiceHorario;
    }

    linha +=
      Math.max(horarios.length, 1) + 1;
  }

  return null;
}


/************************************************************
 * ENCONTRAR LINHA DOS DETALHES
 ************************************************************/

function encontrarLinhaDetalhePorDados_(
  ocorrencias,
  ocorrencia,
  linhaInicial
) {

  const indice =
    ocorrencias.findIndex(function(o) {

      return (
        o.data === ocorrencia.data &&
        o.funcionario === ocorrencia.funcionario &&
        o.registro === ocorrencia.registro
      );

    });

  if (indice < 0) {
    return null;
  }

  return linhaInicial + indice;
}


/************************************************************
 * CLASSIFICAR OCORRÊNCIA
 ************************************************************/

function classificarOcorrencia_(ocorrencia) {

  const texto =
    (
      String(ocorrencia.situacao || '') +
      ' ' +
      String(ocorrencia.motivo || '')
    ).toLowerCase();


  /*
   * ATESTADO SEMPRE AMARELO.
   */
  if (
    String(ocorrencia.situacao || '') ===
    'Atestado'
  ) {

    return {
      tipo: 'atestado',
      bg: APP.CORES.MEDICO_BG,
      text: APP.CORES.MEDICO_TEXT
    };
  }


  const medico =
    /médic|medic|consulta|exame|dentista|hospital|atendimento/
      .test(texto);

  if (
    medico ||
    ocorrencia.situacao === 'Horário diferente'
  ) {

    return {
      tipo: 'medico',
      bg: APP.CORES.MEDICO_BG,
      text: APP.CORES.MEDICO_TEXT
    };

  }


  return {
    tipo: 'ocorrencia',
    bg: APP.CORES.OCORRENCIA_BG,
    text: APP.CORES.OCORRENCIA_TEXT
  };
}


/************************************************************
 * LER OCORRÊNCIAS DA ABA
 ************************************************************/

function lerOcorrenciasAba_(sheet) {

  const valores =
    sheet
      .getDataRange()
      .getDisplayValues();

  let linhaCabecalho = -1;

  for (
    let i = 0;
    i < valores.length;
    i++
  ) {

    if (
      String(valores[i][0]).trim() === 'Data' &&
      String(valores[i][1]).trim() === 'Funcionário'
    ) {

      linhaCabecalho =
        i;

      break;
    }
  }

  if (linhaCabecalho === -1) {
    return [];
  }

  const ocorrencias = [];

  for (
    let i = linhaCabecalho + 1;
    i < valores.length;
    i++
  ) {

    const linha =
      valores[i];

    if (
      !linha[0] &&
      !linha[1]
    ) {
      continue;
    }

    ocorrencias.push({

      data:
        linha[0],

      funcionario:
        linha[1],

      registro:
        linha[2],

      horarioPrevisto:
        linha[3],

      horarioInformado:
        linha[4],

      situacao:
        linha[5],

      motivo:
        linha[6],

      observacao:
        linha[7]

    });
  }

  return ocorrencias;
}


/************************************************************
 * RELATÓRIO
 *
 * Consolida todos os dados das seções:
 * 📌 DETALHES DAS OCORRÊNCIAS
 *
 * REGRA ESPECIAL:
 * Atestado é agrupado por:
 *
 * Funcionário + Data
 *
 * Portanto:
 *
 * Entrada    → Atestado
 * Saída 1    → Atestado
 * Entrada 2  → Atestado
 * Saída 2    → Atestado
 *
 * vira apenas:
 *
 * Dia inteiro → Atestado
 ************************************************************/

function criarRelatorio_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let sheet =
    ss.getSheetByName('RELATÓRIO');

  if (!sheet) {
    sheet =
      ss.insertSheet('RELATÓRIO');
  }

  atualizarRelatorio_();
}


/************************************************************
 * ATUALIZAR RELATÓRIO
 ************************************************************/

function atualizarRelatorio_() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let relatorio =
    ss.getSheetByName('RELATÓRIO');

  if (!relatorio) {
    relatorio =
      ss.insertSheet('RELATÓRIO');
  }

  relatorio.clear();


  /**********************************************************
   * TÍTULO
   **********************************************************/

  relatorio
    .getRange('A1:H1')
    .breakApart()
    .merge()
    .setValue('📊 RELATÓRIO DE OCORRÊNCIAS')
    .setBackground(APP.CORES.CABECALHO)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setFontSize(16)
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');

  relatorio.setRowHeight(1, 32);


  /**********************************************************
   * CABEÇALHO
   **********************************************************/

  relatorio
    .getRange('A3:H3')
    .setValues([[
      'Data',
      'Funcionário',
      'Registro',
      'Horário previsto',
      'Horário informado',
      'Situação',
      'Motivo',
      'Observação'
    ]])
    .setBackground(APP.CORES.CABECALHO_2)
    .setFontColor('#ffffff')
    .setFontWeight('bold')
    .setHorizontalAlignment('center')
    .setVerticalAlignment('middle');


  /**********************************************************
   * LOCALIZAR ABAS MENSAIS
   **********************************************************/

  const abasMensais =
    ss.getSheets().filter(function(sheet) {

      return /^[A-ZÇ]{3}-\d{4}$/.test(
        sheet.getName()
      );

    });


  /**********************************************************
   * COLETAR DADOS
   **********************************************************/

  const ocorrenciasNormais = [];

  /*
   * Objeto utilizado para agrupar os atestados.
   */
  const atestados = {};


  abasMensais.forEach(function(sheet) {

    const ocorrencias =
      lerOcorrenciasAba_(sheet);

    ocorrencias.forEach(function(o) {


      /******************************************************
       * ATESTADO
       ******************************************************/

      if (
        String(o.situacao || '').trim() ===
        'Atestado'
      ) {

        let dataChave;

        try {

          const data =
            converterData_(o.data);

          dataChave =
            Utilities.formatDate(
              data,
              Session.getScriptTimeZone(),
              'yyyy-MM-dd'
            );

        } catch (e) {

          dataChave =
            String(o.data);

        }


        /*
         * Funcionário + Data
         */
        const chave =
          o.funcionario +
          '|' +
          dataChave;


        /*
         * Primeiro registro daquele dia.
         */
        if (!atestados[chave]) {

          atestados[chave] = {

            data:
              o.data,

            funcionario:
              o.funcionario,

            registro:
              'Dia inteiro',

            horarioPrevisto:
              '',

            horarioInformado:
              '',

            situacao:
              'Atestado',

            motivo:
              o.motivo || '',

            observacao:
              o.observacao || ''

          };

        } else {

          /*
           * Se outro registro do mesmo atestado
           * possuir motivo ou observação preenchidos,
           * preservamos a informação.
           */

          if (
            !atestados[chave].motivo &&
            o.motivo
          ) {

            atestados[chave].motivo =
              o.motivo;

          }

          if (
            !atestados[chave].observacao &&
            o.observacao
          ) {

            atestados[chave].observacao =
              o.observacao;

          }

        }

        return;
      }


      /******************************************************
       * OUTRAS OCORRÊNCIAS
       ******************************************************/

      ocorrenciasNormais.push({

        data:
          o.data,

        funcionario:
          o.funcionario,

        registro:
          o.registro,

        horarioPrevisto:
          o.horarioPrevisto,

        horarioInformado:
          o.horarioInformado,

        situacao:
          o.situacao,

        motivo:
          o.motivo,

        observacao:
          o.observacao

      });

    });

  });


  /**********************************************************
   * MONTAR RESULTADO FINAL
   **********************************************************/

  const todosDados = [];


  /*
   * Ocorrências normais continuam individuais.
   */
  ocorrenciasNormais.forEach(function(o) {

    todosDados.push([

      o.data,
      o.funcionario,
      o.registro,
      o.horarioPrevisto,
      o.horarioInformado,
      o.situacao,
      o.motivo,
      o.observacao

    ]);

  });


  /*
   * Atestados entram apenas uma vez.
   */
  Object.keys(atestados).forEach(function(chave) {

    const o =
      atestados[chave];

    todosDados.push([

      o.data,
      o.funcionario,
      o.registro,
      o.horarioPrevisto,
      o.horarioInformado,
      o.situacao,
      o.motivo,
      o.observacao

    ]);

  });


  /**********************************************************
   * ORDENAR POR DATA E FUNCIONÁRIO
   **********************************************************/

  todosDados.sort(function(a, b) {

    const dataA =
      String(a[0]);

    const dataB =
      String(b[0]);

    if (dataA !== dataB) {

      return dataA.localeCompare(
        dataB
      );

    }

    return String(a[1]).localeCompare(
      String(b[1]),
      'pt-BR'
    );

  });


  /**********************************************************
   * ESCREVER DADOS
   **********************************************************/

  if (todosDados.length) {

    relatorio
      .getRange(
        4,
        1,
        todosDados.length,
        8
      )
      .setValues(todosDados)
      .setVerticalAlignment('middle')
      .setWrap(true);


    /******************************************************
     * CORES
     ******************************************************/

    todosDados.forEach(function(
      linha,
      indice
    ) {

      const situacao =
        String(linha[5] || '');


      let background =
        '#ffffff';

      let textColor =
        '#1f2937';


      /*
       * ATESTADO = AMARELO
       */
      if (
        situacao === 'Atestado'
      ) {

        background =
          APP.CORES.MEDICO_BG;

        textColor =
          APP.CORES.MEDICO_TEXT;


      /*
       * HORÁRIO DIFERENTE = AMARELO
       */
      } else if (
        situacao === 'Horário diferente'
      ) {

        background =
          APP.CORES.MEDICO_BG;

        textColor =
          APP.CORES.MEDICO_TEXT;


      /*
       * OUTRAS OCORRÊNCIAS = VERMELHO
       */
      } else {

        background =
          APP.CORES.OCORRENCIA_BG;

        textColor =
          APP.CORES.OCORRENCIA_TEXT;

      }


      relatorio
        .getRange(
          indice + 4,
          1,
          1,
          8
        )
        .setBackground(background)
        .setFontColor(textColor);

    });

  }


  /**********************************************************
   * LARGURAS
   **********************************************************/

  const larguras = [
    100,
    140,
    110,
    130,
    130,
    150,
    260,
    260
  ];


  larguras.forEach(function(
    largura,
    indice
  ) {

    relatorio.setColumnWidth(
      indice + 1,
      largura
    );

  });


  /**********************************************************
   * FINALIZAÇÃO
   **********************************************************/

  relatorio.setFrozenRows(3);
}


/************************************************************
 * VERIFICAR FÉRIAS
 ************************************************************/

function estaDeFerias_(
  funcionario,
  data,
  ferias
) {

  return ferias.some(function(f) {

    if (
      f.funcionario !== funcionario
    ) {
      return false;
    }

    const inicio =
      converterData_(f.inicio);

    const fim =
      converterData_(f.fim);

    inicio.setHours(0, 0, 0, 0);
    fim.setHours(23, 59, 59, 999);

    return (
      data >= inicio &&
      data <= fim
    );

  });
}


/************************************************************
 * COMPARAR DATAS
 ************************************************************/

function mesmaData_(
  valor,
  data
) {

  try {

    const outra =
      converterData_(valor);

    return (
      outra.getFullYear() ===
      data.getFullYear() &&

      outra.getMonth() ===
      data.getMonth() &&

      outra.getDate() ===
      data.getDate()
    );

  } catch (e) {

    return false;
  }
}


/************************************************************
 * ABRIR ABA RELATÓRIO
 ************************************************************/

function abrirAbaRelatorio() {

  const ss =
    SpreadsheetApp.getActiveSpreadsheet();

  let relatorio =
    ss.getSheetByName('RELATÓRIO');

  if (!relatorio) {

    criarRelatorio_();

    relatorio =
      ss.getSheetByName('RELATÓRIO');
  }

  atualizarRelatorio_();

  ss.setActiveSheet(relatorio);

  return 'Relatório atualizado.';
}

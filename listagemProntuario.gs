/*********************************
 * CONFIG
 *********************************/
const CRO_CFG = {
  PASTA_ID: "1Ohu113aiBBCK5FQ-FQDuBpVKa82aB-CO",
  ABA_BASE: "BASE_SCIH_NHE",
  ABA_ANALISE: "ANÁLISE_ÓBITOS",
  ABA_LISTA: "LISTA_SEPARAÇÃO_ÓBITOS",
};

/*********************************
 * ABRE MODAL (BOTÃO)
 *********************************/
function abrirCRO() {
  const html = HtmlService
    .createHtmlOutputFromFile("CRO_HTML")
    .setWidth(980)
    .setHeight(720);

  SpreadsheetApp.getUi().showModalDialog(html, "Separação CRO");
}

/*********************************
 * NORMALIZA DATA
 *********************************/
function normalizarData_(v) {
  if (v instanceof Date) return v;
  if (typeof v === "number" && Number.isFinite(v)) {
    return new Date(Math.round((v - 25569) * 86400 * 1000));
  }
  if (typeof v === "string") {
    const s = v.trim();
    if (!s) return null;
    const cleaned = s.split(" ")[0];
    if (cleaned.includes("/")) {
      const p = cleaned.split("/");
      if (p.length === 3) {
        const dia = Number(p[0]);
        const mes = Number(p[1]) - 1;
        const ano = Number(p[2]);
        if (Number.isFinite(dia) && Number.isFinite(mes) && Number.isFinite(ano)) {
          return new Date(ano, mes, dia);
        }
      }
    }
    if (cleaned.includes("-")) {
      const p = cleaned.split("-");
      if (p.length === 3) {
        if (p[0].length === 4) {
          const ano = Number(p[0]);
          const mes = Number(p[1]) - 1;
          const dia = Number(p[2]);
          if (Number.isFinite(dia) && Number.isFinite(mes) && Number.isFinite(ano)) {
            return new Date(ano, mes, dia);
          }
        } else {
          const dia = Number(p[0]);
          const mes = Number(p[1]) - 1;
          const ano = Number(p[2]);
          if (Number.isFinite(dia) && Number.isFinite(mes) && Number.isFinite(ano)) {
            return new Date(ano, mes, dia);
          }
        }
      }
    }
  }
  return null;
}

// Coluna "DATA ÓBITO" na BASE_SCIH_NHE (S)
const COL_DATA_OBITO_BASE = 18;

function normalizarCabecalho_(valor) {
  return String(valor || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/[^A-Za-z0-9]/g, "")
    .toUpperCase();
}

function obterIndiceCabecalho_(cabecalhos, candidatos) {
  const mapa = new Map(
    cabecalhos.map((c, i) => [normalizarCabecalho_(c), i])
  );
  for (const candidato of candidatos) {
    const idx = mapa.get(normalizarCabecalho_(candidato));
    if (idx !== undefined) return idx;
  }
  return -1;
}

/*********************************
 * ANOS DISPONÍVEIS (BASE col F)
 *********************************/
function obterAnosDisponiveisCRO() {
  const sh = SpreadsheetApp.getActive().getSheetByName(CRO_CFG.ABA_BASE);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const anosColF = sh.getRange(2, 6, lastRow - 1, 1).getValues();
  const set = new Set();

  anosColF.forEach(r => r[0] && set.add(Number(r[0])));
  return Array.from(set).sort((a, b) => a - b);
}

/*********************************
 * ANALISAR ÓBITOS (CONTADORES)
 *********************************/
function analisarObitosCRO(mes, ano) {
  const ss = SpreadsheetApp.getActive();
  const baseData = ss.getSheetByName(CRO_CFG.ABA_BASE).getDataRange().getValues();
  const analiseData = ss.getSheetByName(CRO_CFG.ABA_ANALISE).getDataRange().getValues();

  const baseHeaders = baseData[0] || [];
  const analiseHeaders = analiseData[0] || [];

  const idxDataObito = obterIndiceCabecalho_(baseHeaders, ["DATA ÓBITO", "DATA OBITO", "DATA DO ÓBITO", "DATA DO OBITO"]);
  const idxProntBase = obterIndiceCabecalho_(baseHeaders, ["PRONT", "PRONTUÁRIO", "PRONTUARIO"]);
  const idxProntAnalise = obterIndiceCabecalho_(analiseHeaders, ["PRONT", "PRONTUÁRIO", "PRONTUARIO"]);
  const idxStatusAnalise = obterIndiceCabecalho_(analiseHeaders, ["1ª AVALIAÇÃO CONCLUÍDA?", "1A AVALIACAO CONCLUIDA", "STATUS"]);

  const avaliadosSet = new Set(
    analiseData.slice(1)
      .filter(r =>
        (r[idxStatusAnalise > -1 ? idxStatusAnalise : 6] === "SIM") ||
        (r[idxStatusAnalise > -1 ? idxStatusAnalise : 6] === "AGUARDA PROT. LONDRES") ||
        (r[idxStatusAnalise > -1 ? idxStatusAnalise : 6] === "LONDRES AVALIADO")
      )
      .map(r => String(r[idxProntAnalise > -1 ? idxProntAnalise : 8]).trim())
  );

  let avaliados = 0;
  let pendentes = 0;

  baseData.slice(1).forEach(r => {
    const data = normalizarData_(r[idxDataObito > -1 ? idxDataObito : COL_DATA_OBITO_BASE]);
    if (!data) return;

    if (data.getMonth() + 1 === mes && data.getFullYear() === ano) {
      avaliadosSet.has(String(r[idxProntBase > -1 ? idxProntBase : 8]).trim())
        ? avaliados++
        : pendentes++;
    }
  });

  return { total: avaliados + pendentes, avaliados, pendentes };
}

/*********************************
 * GERAR LISTA + RETORNAR PREVIEW
 *********************************/
function gerarListaCRO(mes, ano) {
  const ss = SpreadsheetApp.getActive();
  const lista = ss.getSheetByName(CRO_CFG.ABA_LISTA);

  lista.getRange("A3:H").clearContent();

  const baseData = ss.getSheetByName(CRO_CFG.ABA_BASE).getDataRange().getValues();
  const analiseData = ss.getSheetByName(CRO_CFG.ABA_ANALISE).getDataRange().getValues();

  const baseHeaders = baseData[0] || [];
  const analiseHeaders = analiseData[0] || [];

  const idxDataObito = obterIndiceCabecalho_(baseHeaders, ["DATA ÓBITO", "DATA OBITO", "DATA DO ÓBITO", "DATA DO OBITO"]);
  const idxProntBase = obterIndiceCabecalho_(baseHeaders, ["PRONT", "PRONTUÁRIO", "PRONTUARIO"]);
  const idxNomeBase = obterIndiceCabecalho_(baseHeaders, ["NOME", "NOME COMPLETO"]);
  const idxUnidadeBase = obterIndiceCabecalho_(baseHeaders, ["UNIDADE DO ÓBITO", "UNIDADE DO OBITO", "UNIDADE"]);
  const idxNascBase = obterIndiceCabecalho_(baseHeaders, ["DN", "DATA NASCIMENTO", "DATA DE NASCIMENTO"]);

  const idxProntAnalise = obterIndiceCabecalho_(analiseHeaders, ["PRONT", "PRONTUÁRIO", "PRONTUARIO"]);
  const idxStatusAnalise = obterIndiceCabecalho_(analiseHeaders, ["1ª AVALIAÇÃO CONCLUÍDA?", "1A AVALIACAO CONCLUIDA", "STATUS"]);
  const idxStatusLista = obterIndiceCabecalho_(analiseHeaders, ["STATUS"]);

  const avaliadosSet = new Set(
    analiseData.slice(1)
      .filter(r =>
        (r[idxStatusAnalise > -1 ? idxStatusAnalise : 6] === "SIM") ||
        (r[idxStatusAnalise > -1 ? idxStatusAnalise : 6] === "AGUARDA PROT. LONDRES") ||
        (r[idxStatusAnalise > -1 ? idxStatusAnalise : 6] === "LONDRES AVALIADO")
      )
      .map(r => String(r[idxProntAnalise > -1 ? idxProntAnalise : 8]).trim())
  );

  const statusMap = new Map(
    analiseData.slice(1).map(r => [
      String(r[idxProntAnalise > -1 ? idxProntAnalise : 8]).trim(),
      r[idxStatusLista > -1 ? idxStatusLista : 3]
    ])
  );

  const saida = [];

  baseData.slice(1).forEach(r => {
    const data = normalizarData_(r[idxDataObito > -1 ? idxDataObito : COL_DATA_OBITO_BASE]);
    if (!data) return;

    if (data.getMonth() + 1 === mes && data.getFullYear() === ano) {
      const pront = String(r[idxProntBase > -1 ? idxProntBase : 8]).trim();
      if (!avaliadosSet.has(pront)) {
        saida.push([
          r[idxProntBase > -1 ? idxProntBase : 8],   // PRONT
          r[idxNomeBase > -1 ? idxNomeBase : 9],     // NOME
          "",     // AVALIADOR
          r[idxUnidadeBase > -1 ? idxUnidadeBase : 57],  // UNIDADE
          r[idxDataObito > -1 ? idxDataObito : COL_DATA_OBITO_BASE],   // DATA ÓBITO
          r[idxNascBase > -1 ? idxNascBase : 10],  // NASC
          statusMap.get(pront) || "",
          ""
        ]);
      }
    }
  });

  if (saida.length) {
    lista.getRange(3, 1, saida.length, 8).setValues(saida);
  }

  SpreadsheetApp.flush(); // 🔴 ESSENCIAL

  const preview = saida.map(linha =>
    linha.map(valor =>
      valor instanceof Date
        ? Utilities.formatDate(valor, Session.getScriptTimeZone(), "dd/MM/yyyy")
        : valor
    )
  );

  // Garantir que o preview reflita exatamente o que foi gerado
  // (evita retorno vazio por problemas de serialização).
  return preview.length ? preview : []; // 🔴 preview vem daqui
}

/*********************************
 * SALVAR OBSERVAÇÃO (H)
 *********************************/
function salvarObservacaoCRO(pront, obs) {
  const sh = SpreadsheetApp.getActive().getSheetByName(CRO_CFG.ABA_LISTA);
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return;

  const range = sh.getRange(3, 1, lastRow - 2, 8);
  const dados = range.getValues();

  for (let i = 0; i < dados.length; i++) {
    if (String(dados[i][0]).trim() === String(pront).trim()) {
      dados[i][7] = obs;
      break;
    }
  }

  range.setValues(dados);
}

/*********************************
 * GERAR PDF (A2 até última)
 *********************************/
function gerarPdfCRO(mes, ano) {
  const ss = SpreadsheetApp.getActive();
  const sh = ss.getSheetByName(CRO_CFG.ABA_LISTA);
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return { ok: false };

  const pasta = DriveApp.getFolderById(CRO_CFG.PASTA_ID);
  const gid = sh.getSheetId();
  const nome = `${String(mes).padStart(2, "0")}_${ano}_CRO.pdf`;

  const url =
    ss.getUrl().replace(/edit$/, "") +
    "export?format=pdf" +
    "&gid=" + gid +
    "&range=A2:H" + lastRow +
    "&portrait=false&fitw=true&sheetnames=false&gridlines=false&pagenumbers=false";

  const blob = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + ScriptApp.getOAuthToken() }
  }).getBlob().setName(nome);

  pasta.createFile(blob);
  return { ok: true };
}

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
  const base = ss.getSheetByName(CRO_CFG.ABA_BASE).getDataRange().getValues();
  const analise = ss.getSheetByName(CRO_CFG.ABA_ANALISE).getDataRange().getValues();

  const avaliadosSet = new Set(
    analise.slice(1)
      .filter(r =>
        r[6] === "SIM" ||
        r[6] === "AGUARDA PROT. LONDRES" ||
        r[6] === "LONDRES AVALIADO"
      )
      .map(r => String(r[8]).trim())
  );

  let avaliados = 0;
  let pendentes = 0;

  base.slice(1).forEach(r => {
    const data = normalizarData_(r[COL_DATA_OBITO_BASE]);
    if (!data) return;

    if (data.getMonth() + 1 === mes && data.getFullYear() === ano) {
      avaliadosSet.has(String(r[8]).trim())
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

  const base = ss.getSheetByName(CRO_CFG.ABA_BASE).getDataRange().getValues();
  const analise = ss.getSheetByName(CRO_CFG.ABA_ANALISE).getDataRange().getValues();

  const avaliadosSet = new Set(
    analise.slice(1)
      .filter(r =>
        r[6] === "SIM" ||
        r[6] === "AGUARDA PROT. LONDRES" ||
        r[6] === "LONDRES AVALIADO"
      )
      .map(r => String(r[8]).trim())
  );

  const statusMap = new Map(
    analise.slice(1).map(r => [String(r[8]).trim(), r[3]])
  );

  const saida = [];

  base.slice(1).forEach(r => {
    const data = normalizarData_(r[COL_DATA_OBITO_BASE]);
    if (!data) return;

    if (data.getMonth() + 1 === mes && data.getFullYear() === ano) {
      const pront = String(r[8]).trim();
      if (!avaliadosSet.has(pront)) {
        saida.push([
          r[8],   // PRONT
          r[9],   // NOME
          "",     // AVALIADOR
          r[57],  // UNIDADE
          r[COL_DATA_OBITO_BASE],   // DATA ÓBITO
          r[10],  // NASC
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

  // Garantir que o preview reflita exatamente o que foi gravado na planilha
  // (evita retornar [] quando o retorno direto não é serializado pelo Apps Script).
  return saida.length
    ? lista.getRange(3, 1, saida.length, 8).getValues()
    : []; // 🔴 preview vem daqui
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

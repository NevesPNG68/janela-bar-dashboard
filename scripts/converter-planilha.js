// Converte a planilha Por-venda_2026.xlsm em um arquivo leve (dados/painel.json.gz)
// para o painel abrir sem precisar baixar e decodificar o Excel inteiro.
//
// Usa a MESMA biblioteca e as MESMAS opções de leitura do painel (SheetJS 0.18.5),
// então os valores entregues ao painel são idênticos aos de hoje.
//
// Uso: node scripts/converter-planilha.js <planilha.xlsm> <saida.json.gz>

const fs = require("fs");
const zlib = require("zlib");
const XLSX = require("xlsx");

const [, , entrada, saida] = process.argv;
if (!entrada || !saida) {
  console.error("Uso: node converter-planilha.js <planilha.xlsm> <saida.json.gz>");
  process.exit(1);
}

// Abas que o painel procura (mesmos termos usados no index.html).
const TERMOS = [
  "venda", "despesa", "matrizanalitico", "matrizanaltico", "projeto", "projecao",
  "atracao", "analisedrink", "anlisedrink", "dadosgrafkpi", "grafkpi", "kpis",
];
const normAmplo = s => String(s || "").trim().toLowerCase()
  .normalize("NFD").replace(/[̀-ͯ]/g, "").replace(/[^a-z0-9]/g, "");
const usada = nome => TERMOS.some(t => normAmplo(nome).includes(t));

// Na aba de vendas (283 mil linhas) o painel só lê estas colunas (índice 0 = coluna A).
// B, G, I, J, M, N, Q até AA. As demais (ex.: códigos de pedido) ficam vazias no arquivo leve.
// Se o painel passar a usar outra coluna da VendasOperador, inclua o índice aqui.
const COLUNAS_VENDAS = new Set([1, 6, 8, 9, 12, 13, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26]);
const ehVendas = nome => normAmplo(nome).includes("venda");

const t0 = Date.now();
const buf = fs.readFileSync(entrada);

// 1ª leitura rápida: só nomes das abas.
const nomes = XLSX.read(buf, { type: "buffer", bookSheets: true }).SheetNames;
const abasUsadas = nomes.filter(usada);
console.log("Abas na planilha:", nomes.length, "| usadas pelo painel:", abasUsadas.join(", "));

// 2ª leitura: só as abas usadas, com as opções do painel.
const wb = XLSX.read(buf, {
  type: "buffer", dense: true, cellFormula: false, cellHTML: false,
  cellStyles: false, cellText: false, sheets: abasUsadas,
});
console.log("Leitura concluída em", ((Date.now() - t0) / 1000).toFixed(1), "s");

// Codificação em colunas: textos repetidos viram índices de um dicionário.
function codificarAba(sheet, colunasMantidas) {
  const linhas = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null });
  const nLin = linhas.length;
  const nCol = linhas.reduce((m, r) => Math.max(m, r.length), 0);
  const colunas = [];
  for (let c = 0; c < nCol; c++) {
    if (colunasMantidas && !colunasMantidas.has(c)) { colunas.push({ t: "v" }); continue; }
    let temTexto = false, temOutro = false;
    for (let r = 0; r < nLin; r++) {
      const v = linhas[r][c];
      if (v === null || v === undefined) continue;
      if (typeof v === "string") temTexto = true; else temOutro = true;
    }
    if (!temTexto && !temOutro) { colunas.push({ t: "v" }); continue; }         // vazia
    if (temTexto && !temOutro) {                                                  // só textos
      const dic = [], idx = new Map(), i = new Array(nLin);
      for (let r = 0; r < nLin; r++) {
        const v = linhas[r][c];
        if (v === null || v === undefined) { i[r] = -1; continue; }
        let k = idx.get(v);
        if (k === undefined) { k = dic.length; dic.push(v); idx.set(v, k); }
        i[r] = k;
      }
      colunas.push({ t: "s", d: dic, i });
      continue;
    }
    const d = new Array(nLin);                                                    // números/misto
    for (let r = 0; r < nLin; r++) { const v = linhas[r][c]; d[r] = v === undefined ? null : v; }
    colunas.push({ t: "m", d });
  }
  return { n: nLin, w: nCol, c: colunas };
}

const saidaObj = {
  versao: 1,
  gerado_em: new Date().toISOString(),
  origem: { bytes: buf.length },
  SheetNames: nomes,
  abas: {},
};
for (const nome of abasUsadas) saidaObj.abas[nome] = codificarAba(wb.Sheets[nome], ehVendas(nome) ? COLUNAS_VENDAS : null);

const json = Buffer.from(JSON.stringify(saidaObj));
const gz = zlib.gzipSync(json, { level: 9 });
fs.writeFileSync(saida, gz);
console.log(
  "Gerado", saida, "| JSON", (json.length / 1e6).toFixed(1), "MB | compactado",
  (gz.length / 1e6).toFixed(2), "MB | total", ((Date.now() - t0) / 1000).toFixed(1), "s"
);

'use strict';

/**
 * Web Worker de ingestão: faz o parse do Excel e o ETL fora da thread principal,
 * para a UI não travar durante a leitura de arquivos grandes (~26MB / 100k+ linhas)
 * e para dar progresso real por etapa em vez de um spinner "cego".
 */

function carregarXLSX() {
  if (typeof XLSX !== 'undefined') return;
  try {
    importScripts('vendor/xlsx.full.min.js');
  } catch {
    importScripts('https://cdn.sheetjs.com/xlsx-0.20.3/package/dist/xlsx.full.min.js');
  }
}

carregarXLSX();
importScripts('etl.js');

self.onmessage = (evento) => {
  const { buffer, overrides } = evento.data;
  try {
    if (typeof XLSX === 'undefined') {
      throw new Error('SheetJS não carregou dentro do worker (vendor/ e CDN indisponíveis).');
    }

    self.postMessage({ tipo: 'progresso', etapa: 'lendo' });
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });

    self.postMessage({ tipo: 'progresso', etapa: 'detectando-aba' });
    const nomeAba = ETL.encontrarAbaAlvo(workbook);
    const worksheet = workbook.Sheets[nomeAba];

    self.postMessage({ tipo: 'progresso', etapa: 'mapeando-colunas' });
    const resultado = ETL.construirRegistros(worksheet, overrides, (atual, total) => {
      self.postMessage({ tipo: 'progresso', etapa: 'processando-linhas', atual, total });
    });

    self.postMessage({ tipo: 'progresso', etapa: 'calculando-indicadores' });
    const kpis = ETL.calcularKPIs(resultado.registros);
    const grupos = ETL.agruparPorSetor(resultado.registros);

    self.postMessage({
      tipo: 'resultado',
      nomeAba,
      registros: resultado.registros,
      mapa: resultado.mapa,
      mapaAutomatico: resultado.mapaAutomatico,
      linhaCabecalho: resultado.linhaCabecalho,
      headers: resultado.headers,
      amostras: resultado.amostras,
      confiancaCabecalho: resultado.confiancaCabecalho,
      kpis,
      grupos,
    });
  } catch (err) {
    self.postMessage({ tipo: 'erro', mensagem: err && err.message ? err.message : String(err) });
  }
};

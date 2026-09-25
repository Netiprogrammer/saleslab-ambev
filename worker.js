'use strict';

/**
 * Web Worker de ingestão: faz o parse do Excel e o ETL fora da thread principal,
 * para a UI não travar durante a leitura de arquivos grandes (~26MB / 100k+ linhas)
 * e para dar progresso real por etapa em vez de um spinner "cego".
 *
 * Lê TODAS as abas relevantes do workbook (não só a maior/primeira) e mescla os
 * registros pela chave do PDV — ver processarWorkbook() em etl.js.
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

    // Arquivos reais de BI corporativo costumam ter dezenas de abas de histórico/rascunho
    // além das poucas que o app usa — testamos restringir a leitura só às abas relevantes
    // (opção `sheets` do SheetJS), mas o custo dominante nesses arquivos é descompactar o
    // .xlsx/.xlsm e a tabela de strings compartilhadas (uma só vez, pro arquivo inteiro),
    // não o parse por aba — então a leitura completa abaixo já é o caminho mais simples
    // e não fica mais lenta por restringir. Em arquivos de dezenas de MB isso pode levar
    // um tempo real (dezenas de segundos); a barra de progresso por etapa existe por isso.
    self.postMessage({ tipo: 'progresso', etapa: 'lendo' });
    const workbook = XLSX.read(buffer, { type: 'array', cellDates: true });

    self.postMessage({ tipo: 'progresso', etapa: 'detectando-abas' });
    const resultado = ETL.processarWorkbook(workbook, overrides || {}, (info) => {
      self.postMessage({
        tipo: 'progresso',
        etapa: 'processando-linhas',
        sheetName: info.sheetName,
        indiceAba: info.indiceAba,
        totalAbas: info.totalAbas,
        atual: info.atual,
        total: info.total,
      });
    });

    self.postMessage({ tipo: 'progresso', etapa: 'calculando-indicadores' });
    const kpis = ETL.calcularKPIs(resultado.registros);
    const grupos = ETL.agruparPorSetor(resultado.registros);

    self.postMessage({
      tipo: 'resultado',
      registros: resultado.registros,
      abasProcessadas: resultado.abasProcessadas,
      pdvsComMultiplasFontes: resultado.pdvsComMultiplasFontes,
      kpis,
      grupos,
    });
  } catch (err) {
    self.postMessage({ tipo: 'erro', mensagem: err && err.message ? err.message : String(err) });
  }
};

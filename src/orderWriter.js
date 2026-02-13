const { writeOrderBlockWithExcelJs } = require('./writers/exceljsWriter');
const { writeOrderBlockWithCom } = require('./writers/comWriter');

async function writeOrderBlock(params) {
  const backend = params.backend || 'exceljs';

  if (backend === 'com') {
    return writeOrderBlockWithCom(params);
  }

  if (backend === 'exceljs') {
    return writeOrderBlockWithExcelJs(params);
  }

  throw new Error(`Unbekanntes writerBackend: ${backend}`);
}

module.exports = {
  writeOrderBlock,
};

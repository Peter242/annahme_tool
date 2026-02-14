const { writeOrderBlockWithExcelJs } = require('./writers/exceljsWriter');
const { writeOrderBlockWithCom } = require('./writers/comWriter');

async function writeOrderBlock(params) {
  const backend = String(params.backend || 'exceljs').toLowerCase();

  if (backend === 'com') {
    return writeOrderBlockWithCom(params);
  }

  if (backend === 'exceljs' || backend === 'comexceljs') {
    return writeOrderBlockWithExcelJs(params);
  }

  throw new Error(`Unbekanntes writerBackend: ${backend}`);
}

module.exports = {
  writeOrderBlock,
};

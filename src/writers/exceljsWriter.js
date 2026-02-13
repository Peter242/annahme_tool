const { appendOrderBlockToYearSheet } = require('../excelCommit');

async function writeOrderBlockWithExcelJs(params) {
  return appendOrderBlockToYearSheet(params);
}

module.exports = {
  writeOrderBlockWithExcelJs,
};

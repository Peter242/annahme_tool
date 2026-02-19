function createImportPackagesHandler(deps) {
  const {
    getConfig,
    resolveExcelPath,
    invalidatePackagesCache,
    importPackagesFromExcel,
    writePackages,
    readPackages,
  } = deps;

  return async function importPackagesHandler(_req, res) {
    try {
      const config = getConfig();
      const excelPath = resolveExcelPath(config.excelPath);
      invalidatePackagesCache();
      const packages = await importPackagesFromExcel(excelPath, 'Vorlagen');
      writePackages(packages);
      invalidatePackagesCache();
      const freshPackages = readPackages({ forceReload: true });
      res.set('Cache-Control', 'no-store');
      return res.json({
        ok: true,
        count: freshPackages.length,
        packages: freshPackages,
      });
    } catch (error) {
      return res.status(400).json({
        ok: false,
        message: error.message,
      });
    }
  };
}

module.exports = {
  createImportPackagesHandler,
};

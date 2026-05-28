const INVENTORY_SCAN_EMERGENCY_CELL_LIMIT = 5000000;

function inventoryScanErrorMessage(error) {
  const message = error?.message || error;
  return String(message || "неизвестная ошибка").replace(/\s+/g, " ").trim();
}

function formatInventorySheetDimensions(sheet) {
  if (sheet.rowCount && sheet.columnCount) {
    return `${sheet.rowCount} строк, ${sheet.columnCount} колонок`;
  }

  if (sheet.rowCount) {
    return `${sheet.rowCount} строк`;
  }

  if (sheet.columnCount) {
    return `${sheet.columnCount} колонок`;
  }

  return "";
}

function formatInventorySkippedSheetLine(sheet) {
  if (sheet.reason === "empty") {
    return `${sheet.sheetName}: пустой лист.`;
  }

  if (sheet.reason === "emergencyLimit") {
    return `${sheet.sheetName}: аварийно пропущен — UsedRange слишком большой (${sheet.rowCount} стр. × ${sheet.columnCount} кол. = ${sheet.cellCount} ячеек, аварийный лимит: ${INVENTORY_SCAN_EMERGENCY_CELL_LIMIT}).`;
  }

  if (sheet.reason === "scanError") {
    return `${sheet.sheetName}: ошибка сканирования — ${sheet.message || "неизвестная ошибка"}.`;
  }

  return `${sheet.sheetName}: пропущен (${sheet.reason || "неизвестная причина"}).`;
}

function buildInventoryContentSkippedRow(sheet) {
  if (sheet.reason === "empty") {
    return [sheet.sheetName, "Skipped", "Пустой лист", "", ""];
  }

  if (sheet.reason === "emergencyLimit") {
    return [
      sheet.sheetName,
      "Skipped",
      "Аварийный стоп сканирования",
      `${sheet.rowCount} строк, ${sheet.columnCount} колонок`,
      `${sheet.cellCount} ячеек; аварийный лимит ${INVENTORY_SCAN_EMERGENCY_CELL_LIMIT}`,
    ];
  }

  if (sheet.reason === "scanError") {
    return [
      sheet.sheetName,
      "Skipped",
      "Ошибка сканирования листа",
      formatInventorySheetDimensions(sheet),
      sheet.message || "неизвестная ошибка",
    ];
  }

  return [
    sheet.sheetName,
    "Skipped",
    "Пропущен",
    "",
    sheet.reason || "неизвестная причина",
  ];
}

export {
  INVENTORY_SCAN_EMERGENCY_CELL_LIMIT,
  inventoryScanErrorMessage,
  formatInventorySkippedSheetLine,
  buildInventoryContentSkippedRow,
};

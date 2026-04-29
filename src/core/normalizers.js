/**
 * Утилита для безопасного извлечения числа из сырого значения ячейки.
 * Берет на себя всю грязную работу с запятыми, пробелами и символами процента.
 */
function parseRawCellValue(rawValue) {
  if (rawValue === null || rawValue === undefined || rawValue === "") {
    return null;
  }

  const textValue = String(rawValue).trim();
  const isPercentText = textValue.endsWith("%");

  const cleanedTextValue = isPercentText
    ? textValue.replace("%", "").trim()
    : textValue;

  const numericValue = Number(cleanedTextValue.replace(",", "."));

  if (Number.isNaN(numericValue)) {
    return null;
  }

  return { numericValue, isPercentText };
}

/**
 * Нормализует доли (проценты).
 * 42% -> 0.42 | 42 -> 0.42 | 0.42 -> 0.42
 */
export function normalizeShare(rawValue) {
  const parsed = parseRawCellValue(rawValue);
  if (!parsed) return null;

  if (parsed.isPercentText) {
    return parsed.numericValue / 100;
  }

  return parsed.numericValue > 1 ? parsed.numericValue / 100 : parsed.numericValue;
}

/**
 * Нормализует стандартное отклонение или дисперсию.
 * Всегда возвращает дисперсию для t-теста Уэлча.
 */
export function normalizeVariance(spreadRawValue, spreadType) {
  const parsed = parseRawCellValue(spreadRawValue);
  if (!parsed || parsed.numericValue < 0) return null;

  const val = parsed.numericValue;
  return spreadType === "standardDeviation" ? val * val : val;
}

/**
 * Нормализует значение NPS (приводит к шкале -1..1).
 * 40 -> 0.4 | -50 -> -0.5
 */
export function normalizeNpsValue(rawValue) {
  const parsed = parseRawCellValue(rawValue);
  if (!parsed) return null;

  const val = parsed.numericValue;
  return Math.abs(val) > 1 ? val / 100 : val;
}

/**
 * Нормализует разброс для NPS (приводит к масштабу долей).
 */
export function normalizeNpsSpread(rawSpread, spreadType) {
  const parsed = parseRawCellValue(rawSpread);
  if (!parsed || parsed.numericValue < 0) return null;

  const val = parsed.numericValue;
  
  if (spreadType === "standardDeviation") {
    return val > 1 ? val / 100 : val;
  }

  if (spreadType === "variance") {
    return val > 1 ? val / 10000 : val;
  }

  return null;
}
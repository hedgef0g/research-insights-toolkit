/**
 * Словарь ключевых слов для автоматического определения типов строк.
 * * ПРАВИЛА ДОБАВЛЕНИЯ:
 * - Все слова должны быть в нижнем регистре.
 * - Без знаков препинания (точки, запятые, скобки).
 * - Вместо буквы "ё" используется "е".
 * * Порядок важен: алгоритм проверяет совпадения сверху вниз.
 */
export const METRIC_DICTIONARY = [
  {
    rowType: "nps",
    keywords: ["nps", "нпс", "net promoter score", "индекс лояльности"],
  },
  {
    rowType: "promoters",
    keywords: ["promoters", "promoter", "промоутеры", "промоутер", "сторонники", "лояльные"],
  },
  {
    rowType: "detractors",
    keywords: ["detractors", "detractor", "детракторы", "детрактор", "критики", "недовольные"],
  },
  {
    rowType: "standardDeviation",
    keywords: [
      "sd", "std", "std dev", "standard deviation", "stdev", "st dev",
      "стандартное отклонение", "ст отклонение", "ст откл", 
      "среднеквадратическое отклонение", "ско",
    ],
  },
  {
    rowType: "variance",
    keywords: ["variance", "var", "dispersion", "дисперсия", "дисп"],
  },
  {
    rowType: "mean",
    keywords: [
      "mean", "average", "avg", "среднее", "средняя", "срзнач", 
      "ср знач", "ср значение", "среднее значение",
    ],
  },
  {
    rowType: "base",
    keywords: ["base", "база", "основание", "выборка", "количество", "кол во"],
  },
  {
    rowType: "proportion",
    keywords: ["%", "percent", "percentage", "share", "доля", "процент", "проценты"],
  },
];
/**
 * Сервис работы со справочниками
 */
let dictionaryCache = null;
let dictionaryCacheTimestamp = 0;
const DICTIONARY_CACHE_KEY = "queue:dictionaries:v1";

class DictionaryService {

  /**
   * Возвращает все справочники
   */
  static getAll() {

    // За время одного построения очереди этот метод вызывается много раз
    // (для резервов и автоназначения). Не перечитываем Google Таблицу для
    // каждой строки, но сохраняем почти мгновенное обновление справочника.
    if (dictionaryCache && Date.now() - dictionaryCacheTimestamp < 60000) {
      return dictionaryCache;
    }

    const cached = CacheService.getScriptCache().get(DICTIONARY_CACHE_KEY);
    if (cached) {
      try {
        dictionaryCache = JSON.parse(cached);
        dictionaryCacheTimestamp = Date.now();
        return dictionaryCache;
      } catch (error) {
        CacheService.getScriptCache().remove(DICTIONARY_CACHE_KEY);
      }
    }

    this.ensureRoleDictionaries();

    const rows = SheetService.getData(
    CONFIG.SHEETS.DICTIONARIES
);

    if (rows.length < 2) {
      return {};
    }

    const headers = rows[0];

    const dictionaries = {};

    headers.forEach(name => {

      if (name) {
        dictionaries[name] = [];
      }

    });

    for (let r = 1; r < rows.length; r++) {

      headers.forEach((name, c) => {

        if (!name) return;

        const value = rows[r][c];

        if (
          value !== "" &&
          value !== null &&
          value !== undefined
        ) {
          dictionaries[name].push(value);
        }

      });

    }

    dictionaryCache = dictionaries;
    dictionaryCacheTimestamp = Date.now();
    CacheService.getScriptCache().put(DICTIONARY_CACHE_KEY, JSON.stringify(dictionaries), 60);
    return dictionaryCache;

  }

  static invalidateCache() {
    dictionaryCache = null;
    dictionaryCacheTimestamp = 0;
    CacheService.getScriptCache().remove(DICTIONARY_CACHE_KEY);
  }

  static ensureRoleDictionaries() {
    SheetService.ensureDictionaryColumn("Конструкторы", ["Валентина", "Ирина"]);
    SheetService.ensureDictionaryColumn("Шлифовка", ["Амир", "Сергей"]);
    SheetService.ensureDictionaryColumn("Покраска", ["Денис"]);
    SheetService.syncDictionaryColumn("Этапы", [
      "ТЗ",
      "Технолог",
      "Проформа",
      "Согласование",
      "Правки",
      "КД",
      "Запуск",
      "Распил",
      "Распил ЛДСП",
      "Распил МДФ",
      "Распил Шпон",
      "Кромка",
      "Кромка Шпон",
      "Присадка",
      "Присадка Шпон",
      "Фрезеровка",
      "Сборка",
      "В покраску",
      "Подготовка",
      "Грунт 1 слой",
      "Шлифовка 1 слой",
      "Грунт 2 слой",
      "Шлифовка 2 слой",
      "Покраска тыльной стороны",
      "Обклейка",
      "Покраска лицевой стороны",
      "Шлифовка перед лаком",
      "Лак",
      "Полировка",
      "Упаковка", "Упаковка Корпуса", "Упаковка Фасада", "СГП", "СГП Корпуса", "СГП Фасада", "Монтаж", "Рекламация", "Завершен"
    ]);
    SheetService.ensureOrderStageValidation();
  }
}

/**
 * Централизованные нормы и почасовой режим планирования Малярки.
 *
 * Лист хранится в нормализованном виде: настройки, профили, операции и нормы
 * находятся в одной таблице, но имеют разные значения в колонке «Тип записи».
 * Стабильные коды операций после создания не меняются.
 */
const PAINTING_NORM_CACHE_KEY = "queue:painting-norms:v1";

class PaintingNormService {

  static invalidateCache() {
    this._config = null;
    CacheService.getScriptCache().remove(PAINTING_NORM_CACHE_KEY);
  }

  static getHeaders() {
    return [
      "Тип записи", "Код", "Название", "Профиль", "Операция", "Роль",
      "Порядок", "Группа", "Типы фасада", "Операции маршрута", "Активна", "Фиксировано, мин",
      "На м², мин", "База площади", "Значение", "Комментарий"
    ];
  }

  static getLegacyOperations() {
    return [
      { stage: "Подготовка", taskType: "preparation", role: "sanding", order: 10, group: "main", active: true, daysKey: "preparation", column: "Дни Подготовка" },
      { stage: "Грунт 1 слой", taskType: "primer1", role: "painting", order: 20, group: "main", active: true, daysKey: "primer1", column: "Дни Грунт 1 слой" },
      { stage: "Шлифовка 1 слой", taskType: "sanding1", role: "sanding", order: 30, group: "main", active: true, daysKey: "sanding1", column: "Дни Шлифовка" },
      { stage: "Грунт 2 слой", taskType: "primer2", role: "painting", order: 40, group: "main", active: true, daysKey: "primer2", column: "Дни Грунт 2 слой", exclusionColumn: "Исключить Грунт 2 слой" },
      { stage: "Шлифовка 2 слой", taskType: "sanding2", role: "sanding", order: 50, group: "main", active: true, daysKey: "sanding2", column: "Дни Шлифовка 2 слой", exclusionColumn: "Исключить Шлифовка 2 слой" },
      { stage: "Покраска тыльной стороны", taskType: "backPainting", role: "painting", order: 60, group: "main", active: true, daysKey: "backPainting", column: "Дни Покраска тыльной стороны", exclusionColumn: "Исключить Покраска тыльной стороны" },
      { stage: "Обклейка", taskType: "masking", role: "sanding", order: 70, group: "main", active: true, daysKey: "masking", column: "Дни Обклейка", exclusionColumn: "Исключить Обклейка" },
      { stage: "Покраска лицевой стороны", taskType: "frontPainting", role: "painting", order: 80, group: "main", active: true, daysKey: "frontPainting", column: "Дни Покраска лицевой стороны" },
      { stage: "Шлифовка перед лаком", taskType: "preLacquerSanding", role: "sanding", order: 90, group: "lacquer", active: true, daysKey: "preLacquerSanding", column: "Дни Шлифовка перед лаком", optionalGroup: "lacquer" },
      { stage: "Лак", taskType: "lacquer", role: "painting", order: 100, group: "lacquer", active: true, daysKey: "lacquer", column: "Дни Лак", optionalGroup: "lacquer" },
      { stage: "Полировка", taskType: "polishing", role: "painting", order: 110, group: "lacquer", active: true, daysKey: "polishing", column: "Дни Полировка", optionalGroup: "lacquer" }
    ];
  }

  static getSeedProfiles() {
    return [
      {
        code: "smooth_mdf", name: "Гладкий МДФ", facadeTypes: ["Гладкий"], active: true,
        routeOperations: ["preparation", "primer1", "sanding1", "primer2", "sanding2", "backPainting", "masking", "frontPainting"]
      },
      {
        code: "milled_standard", name: "Стандартная фрезеровка", facadeTypes: ["Фрезеровка", "Лето", "Зима"], active: true,
        routeOperations: ["preparation", "primer1", "sanding1", "primer2", "sanding2", "backPainting", "masking", "frontPainting"]
      },
      {
        code: "wave_standard", name: "Сложная фрезеровка", facadeTypes: ["Сложная фрезеровка", "Волна"], active: true,
        routeOperations: ["preparation", "primer1", "sanding1", "primer2", "sanding2", "backPainting", "masking", "frontPainting"]
      },
      {
        code: "veneer_matte", name: "Шпон, прозрачный матовый лак", facadeTypes: ["Шпон"], active: true,
        routeOperations: ["preparation", "primer1", "sanding1", "primer2", "sanding2", "lacquer"]
      },
      {
        code: "gloss_smooth", name: "Гладкий МДФ, глянцевый лак", facadeTypes: ["Глянец"], active: true,
        routeOperations: ["preparation", "primer1", "sanding1", "primer2", "sanding2", "backPainting", "masking", "frontPainting", "preLacquerSanding", "lacquer", "polishing"]
      },
      {
        code: "veneer_gloss", name: "Шпон, прозрачный глянцевый лак", facadeTypes: ["Шпон глянец"], active: true,
        routeOperations: ["preparation", "primer1", "sanding1", "primer2", "sanding2", "lacquer", "polishing"]
      }
    ];
  }

  static getSeedNorms() {
    const values = {
      smooth_mdf: {
        preparation: [16, 6.0], primer1: [12, 3.8], sanding1: [6, 5.0],
        primer2: [10, 3.5], sanding2: [6, 5.5], backPainting: [10, 3.2],
        masking: [8, 2.5], frontPainting: [14, 4.0]
      },
      milled_standard: {
        preparation: [20, 8.6], primer1: [14, 4.8], sanding1: [9, 8.0],
        primer2: [12, 4.5], sanding2: [9, 9.0], backPainting: [10, 3.2],
        masking: [8, 2.5], frontPainting: [16, 5.2]
      },
      wave_standard: {
        preparation: [24, 11.7], primer1: [16, 6.3], sanding1: [12, 11.5],
        primer2: [14, 6.0], sanding2: [12, 13.0], backPainting: [10, 3.2],
        masking: [8, 2.5], frontPainting: [18, 7.0]
      },
      veneer_matte: {
        preparation: [20, 8.5], primer1: [12, 4.2], sanding1: [8, 6.5],
        primer2: [10, 4.0], sanding2: [8, 6.8], lacquer: [14, 4.5]
      },
      gloss_smooth: {
        preparation: [16, 6.0], primer1: [12, 3.8], sanding1: [6, 5.0],
        primer2: [10, 3.5], sanding2: [6, 5.5], backPainting: [10, 3.2],
        masking: [8, 2.5], frontPainting: [14, 4.0], preLacquerSanding: [8, 7.0],
        lacquer: [18, 7.5], polishing: [28, 27.2]
      },
      veneer_gloss: {
        preparation: [20, 8.5], primer1: [12, 4.2], sanding1: [8, 6.5],
        primer2: [10, 4.0], sanding2: [8, 6.8], lacquer: [18, 7.5],
        polishing: [28, 27.2]
      }
    };
    const rows = [];
    this.getSeedProfiles().forEach(profile => {
      this.getLegacyOperations().forEach(operation => {
        const norm = values[profile.code][operation.taskType] || ["", ""];
        const areaBasis = operation.taskType === "backPainting" || operation.taskType === "masking"
          ? "back"
          : operation.taskType === "frontPainting" ||
              (profile.code === "gloss_smooth" && ["preLacquerSanding", "lacquer"].includes(operation.taskType))
            ? "front"
            : ["gloss_smooth", "veneer_gloss"].includes(profile.code) && operation.taskType === "polishing"
              ? "polishing"
              : "painting";
        rows.push({
          profile: profile.code,
          operation: operation.taskType,
          fixedMinutes: norm[0],
          minutesPerSqm: norm[1],
          areaBasis,
          comment: operation.taskType === "frontPainting" ? "Окраска лицевой стороны и торцов." :
            (profile.code === "veneer_matte" && operation.taskType === "lacquer"
              ? "Матовый прозрачный лак по шпону без тонировки."
              : profile.code === "gloss_smooth" && operation.taskType === "preLacquerSanding"
                ? "Матирование лицевого цветного слоя перед лаком; рассчитывается по площади лица и торцов."
                : ["gloss_smooth", "veneer_gloss"].includes(profile.code) && operation.taskType === "polishing"
                  ? "Шлифовка лака, полировка и финальный контроль объединены: 28 + 27,2 × P."
                  : "")
        });
      });
    });
    return rows;
  }

  static ensureSheet() {
    const spreadsheet = SheetService.getSpreadsheet();
    const name = CONFIG.SHEETS.PAINTING_NORMS;
    let sheet = spreadsheet.getSheetByName(name);
    let created = false;
    if (!sheet) {
      sheet = spreadsheet.insertSheet(name);
      created = true;
    }
    if (!SheetService._sheets) SheetService._sheets = {};
    SheetService._sheets[name] = sheet;

    const headers = this.getHeaders();
    if (sheet.getLastRow() === 0 || sheet.getLastColumn() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      created = true;
    } else {
      const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
        .map(value => String(value || "").trim());
      const missing = headers.filter(header => !existingHeaders.includes(header));
      if (missing.length) {
        sheet.getRange(1, sheet.getLastColumn() + 1, 1, missing.length).setValues([missing]);
      }
    }
    this.seedMissingRows(sheet);
    if (created) this.formatSheet(sheet);
    return sheet;
  }

  static formatSheet(sheet) {
    sheet.getRange(1, 1, 1, sheet.getLastColumn())
      .setFontWeight("bold")
      .setBackground("#e8eef7");
    sheet.autoResizeColumns(1, sheet.getLastColumn());
    const columns = SheetService.getColumns(sheet);
    ["Фиксировано, мин", "На м², мин"].forEach(name => {
      if (columns[name]) sheet.getRange(2, columns[name], Math.max(sheet.getMaxRows() - 1, 1), 1).setNumberFormat("0.0");
    });
  }

  static seedMissingRows(sheet) {
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(value => String(value || "").trim());
    const index = headers.reduce((result, header, position) => {
      result[header] = position;
      return result;
    }, {});
    const existing = new Set(rows.slice(1).map(row => {
      const type = String(row[index["Тип записи"]] || "").trim();
      if (type === "NORM") return [type, row[index["Профиль"]], row[index["Операция"]]].join("|");
      return [type, row[index["Код"]]].join("|");
    }));
    if (this.migrateSeedRows(rows, index)) {
      sheet.getRange(1, 1, rows.length, headers.length).setValues(rows);
    }
    const append = [];
    const build = values => headers.map(header => values[header] === undefined ? "" : values[header]);
    [
      { code: "schema_version", value: "1.3", comment: "Версия структуры листа." },
      { code: "hourly_planning_enabled", value: "false", comment: "Включается только после успешной проверки норм и активных заказов." }
    ].forEach(setting => {
      const key = "SETTING|" + setting.code;
      if (!existing.has(key)) append.push(build({ "Тип записи": "SETTING", "Код": setting.code, "Значение": setting.value, "Комментарий": setting.comment }));
    });
    this.getSeedProfiles().forEach(profile => {
      const key = "PROFILE|" + profile.code;
      if (!existing.has(key)) append.push(build({
        "Тип записи": "PROFILE", "Код": profile.code, "Название": profile.name,
        "Типы фасада": profile.facadeTypes.join(", "),
        "Операции маршрута": (profile.routeOperations || []).join(", "), "Активна": profile.active
      }));
    });
    this.getLegacyOperations().forEach(operation => {
      const key = "OPERATION|" + operation.taskType;
      if (!existing.has(key)) append.push(build({
        "Тип записи": "OPERATION", "Код": operation.taskType, "Название": operation.stage,
        "Роль": operation.role === "sanding" ? "Шлифовка" : "Покраска",
        "Порядок": operation.order, "Группа": operation.group, "Активна": operation.active
      }));
    });
    this.getSeedNorms().forEach(norm => {
      const key = ["NORM", norm.profile, norm.operation].join("|");
      if (!existing.has(key)) append.push(build({
        "Тип записи": "NORM", "Профиль": norm.profile, "Операция": norm.operation,
        "Фиксировано, мин": norm.fixedMinutes, "На м², мин": norm.minutesPerSqm,
        "База площади": norm.areaBasis || "painting", "Комментарий": norm.comment
      }));
    });
    if (append.length) sheet.getRange(sheet.getLastRow() + 1, 1, append.length, headers.length).setValues(append);
  }

  static migrateSeedRows(rows, index) {
    const routeTargets = this.getSeedProfiles().reduce((result, profile) => {
      result[profile.code] = (profile.routeOperations || []).join(", ");
      return result;
    }, {});
    const oldRoutes = {
      gloss_smooth: "preparation, primer1, sanding1, primer2, sanding2, lacquer, preLacquerSanding, polishing",
      veneer_gloss: "preparation, primer1, sanding1, primer2, sanding2, lacquer, preLacquerSanding, polishing"
    };
    const targets = this.getSeedNorms().reduce((result, norm) => {
      result[norm.profile + "|" + norm.operation] = norm;
      return result;
    }, {});
    const changedPairs = {
      "gloss_smooth|preLacquerSanding": [10, 12],
      "gloss_smooth|polishing": [18, 15.2],
      "veneer_gloss|polishing": [18, 15.2]
    };
    const retiredSeedPairs = {
      "veneer_gloss|preLacquerSanding": [10, 12]
    };
    let changed = false;
    rows.slice(1).forEach(row => {
      const type = String(row[index["Тип записи"]] || "").trim().toUpperCase();
      if (type === "SETTING" && String(row[index["Код"]] || "").trim() === "schema_version" &&
          String(row[index["Значение"]] || "").trim() !== "1.3") {
        row[index["Значение"]] = "1.3";
        changed = true;
        return;
      }
      if (type === "PROFILE") {
        const code = String(row[index["Код"]] || "").trim();
        const current = String(row[index["Операции маршрута"]] || "").trim();
        const target = routeTargets[code];
        if (code === "wave_standard") {
          const facadeTypes = String(row[index["Типы фасада"]] || "").split(",")
            .map(value => value.trim()).filter(Boolean);
          if (facadeTypes.includes("Сложная фрезеровка") && !facadeTypes.includes("Волна")) {
            facadeTypes.push("Волна");
            row[index["Типы фасада"]] = facadeTypes.join(", ");
            changed = true;
          }
        }
        if (target && (!current || current === oldRoutes[code]) && current !== target) {
          row[index["Операции маршрута"]] = target;
          changed = true;
        }
        return;
      }
      if (type !== "NORM") return;
      const key = String(row[index["Профиль"]] || "").trim() + "|" + String(row[index["Операция"]] || "").trim();
      const retiredPair = retiredSeedPairs[key];
      const fixed = this.toNullableNumber(row[index["Фиксировано, мин"]]);
      const per = this.toNullableNumber(row[index["На м², мин"]]);
      if (retiredPair && fixed === retiredPair[0] && per === retiredPair[1]) {
        row[index["Фиксировано, мин"]] = "";
        row[index["На м², мин"]] = "";
        row[index["База площади"]] = "painting";
        row[index["Комментарий"]] = "";
        changed = true;
        return;
      }
      const target = targets[key];
      if (!target || target.fixedMinutes === "" || target.minutesPerSqm === "") return;
      const wasBlank = fixed === null && per === null;
      const oldPair = changedPairs[key];
      const isOldSeed = Boolean(oldPair && fixed === oldPair[0] && per === oldPair[1]);
      if (wasBlank || isOldSeed) {
        row[index["Фиксировано, мин"]] = target.fixedMinutes;
        row[index["На м², мин"]] = target.minutesPerSqm;
        changed = true;
      }
      const currentBasis = String(row[index["База площади"]] || "").trim();
      const basisCanMigrate = wasBlank || isOldSeed ||
        key === "gloss_smooth|lacquer" && fixed === 18 && per === 7.5 && currentBasis === "painting";
      if (basisCanMigrate && currentBasis !== target.areaBasis) {
        row[index["База площади"]] = target.areaBasis;
        changed = true;
      }
      const currentComment = String(row[index["Комментарий"]] || "").trim();
      if (target.comment && (!currentComment || currentComment.indexOf("Справочно: двухслойная эмаль") === 0) && currentComment !== target.comment) {
        row[index["Комментарий"]] = target.comment;
        changed = true;
      }
    });
    return changed;
  }

  static readConfig() {
    if (this._config) return this._config;
    const sheet = this.ensureSheet();
    const cached = CacheService.getScriptCache().get(PAINTING_NORM_CACHE_KEY);
    if (cached) {
      try {
        this._config = JSON.parse(cached);
        return this._config;
      } catch (error) {
        CacheService.getScriptCache().remove(PAINTING_NORM_CACHE_KEY);
      }
    }
    const rows = SheetService.objectsFromRows(sheet.getDataRange().getValues());
    const settings = {};
    const profiles = [];
    const operations = [];
    const norms = [];
    rows.forEach(row => {
      const type = String(row["Тип записи"] || "").trim().toUpperCase();
      if (type === "SETTING") settings[String(row["Код"] || "").trim()] = String(row["Значение"] || "").trim();
      if (type === "PROFILE") profiles.push({
        code: String(row["Код"] || "").trim(),
        name: String(row["Название"] || "").trim(),
        facadeTypes: String(row["Типы фасада"] || "").split(",").map(value => value.trim()).filter(Boolean),
        routeOperations: String(row["Операции маршрута"] || "").split(",").map(value => value.trim()).filter(Boolean),
        active: this.toBoolean(row["Активна"])
      });
      if (type === "OPERATION") operations.push({
        taskType: String(row["Код"] || "").trim(),
        stage: String(row["Название"] || "").trim(),
        role: String(row["Роль"] || "").trim().toLowerCase() === "шлифовка" ? "sanding" : "painting",
        order: Number(row["Порядок"]) || 0,
        group: String(row["Группа"] || "main").trim() || "main",
        active: this.toBoolean(row["Активна"])
      });
      if (type === "NORM") norms.push({
        profile: String(row["Профиль"] || "").trim(),
        operation: String(row["Операция"] || "").trim(),
        fixedMinutes: this.toNullableNumber(row["Фиксировано, мин"]),
        minutesPerSqm: this.toNullableNumber(row["На м², мин"]),
        areaBasis: this.normalizeAreaBasis(row["База площади"]),
        comment: String(row["Комментарий"] || "").trim()
      });
    });
    const legacyByCode = this.getLegacyOperations().reduce((result, operation) => {
      result[operation.taskType] = operation;
      return result;
    }, {});
    operations.forEach(operation => {
      const legacy = legacyByCode[operation.taskType] || {};
      operation.daysKey = legacy.daysKey || operation.taskType;
      operation.column = legacy.column || "";
      operation.exclusionColumn = legacy.exclusionColumn || "";
      operation.optionalGroup = operation.group === "lacquer" ? "lacquer" : "";
    });
    operations.sort((left, right) => left.order - right.order || left.stage.localeCompare(right.stage, "ru"));
    this._config = {
      schemaVersion: settings.schema_version || "1",
      hourlyEnabled: String(settings.hourly_planning_enabled || "false").toLowerCase() === "true",
      profiles,
      operations,
      norms
    };
    CacheService.getScriptCache().put(PAINTING_NORM_CACHE_KEY, JSON.stringify(this._config), 60);
    return this._config;
  }

  static getClientConfig() {
    const config = this.readConfig();
    const validation = this.validateConfiguration(config);
    return {
      hourlyEnabled: config.hourlyEnabled,
      profiles: config.profiles,
      operations: config.operations.map(operation => ({
        taskType: operation.taskType, stage: operation.stage, role: operation.role,
        order: operation.order, group: operation.group, active: operation.active
      })),
      norms: config.norms,
      validation
    };
  }

  static isHourlyEnabled() {
    return this.readConfig().hourlyEnabled;
  }

  static getOperations(activeOnly) {
    const operations = this.readConfig().operations;
    return activeOnly === false ? operations : operations.filter(operation => operation.active);
  }

  static getProfileForFacadeType(facadeType, config) {
    const value = String(facadeType || "").trim();
    const source = config || this.readConfig();
    return source.profiles.find(profile => profile.active && profile.facadeTypes.includes(value)) || null;
  }

  static calculateOperation(order, operation, config) {
    const source = config || this.readConfig();
    const profile = this.getProfileForFacadeType(order && order.facadeType, source);
    if (!profile) return { valid: false, message: "Для типа фасада не назначен нормативный профиль." };
    const norm = source.norms.find(item => item.profile === profile.code && item.operation === operation.taskType);
    if (!norm || norm.fixedMinutes === null || norm.minutesPerSqm === null) {
      return { valid: false, message: `Не заполнена норма «${operation.stage}» для профиля «${profile.name}».` };
    }
    const areaBasis = this.normalizeAreaBasis(norm.areaBasis);
    const areaFields = { painting: "paintingArea", back: "backPaintingArea", front: "frontPaintingArea", polishing: "polishingArea" };
    const area = Number(order && order[areaFields[areaBasis]]);
    if (!(area > 0)) {
      const labels = {
        painting: "Не указана общая площадь грунтования.",
        back: "Не указана площадь тыльной стороны.",
        front: "Не указана площадь лица и торцов.",
        polishing: "Не указана площадь полировки."
      };
      return { valid: false, message: labels[areaBasis] };
    }
    const rawMinutes = Math.round((norm.fixedMinutes + norm.minutesPerSqm * area) * 10) / 10;
    const roundedMinutes = Math.ceil(rawMinutes / 5) * 5;
    return {
      valid: true,
      profileCode: profile.code,
      profileName: profile.name,
      fixedMinutes: norm.fixedMinutes,
      minutesPerSqm: norm.minutesPerSqm,
      areaBasis,
      area,
      rawMinutes,
      roundedMinutes,
      normHours: Math.round(rawMinutes / 6) / 10
    };
  }

  static previewOrder(order) {
    const config = this.readConfig();
    const operations = this.getActiveOperationsForOrder(order, config);
    const rows = operations.map(operation => Object.assign({
      taskType: operation.taskType,
      stage: operation.stage,
      role: operation.role
    }, this.calculateOperation(order, operation, config)));
    return {
      hourlyEnabled: config.hourlyEnabled,
      profile: this.getProfileForFacadeType(order && order.facadeType, config),
      operations: rows,
      totalRawMinutes: Math.round(rows.reduce((sum, row) => sum + (row.valid ? row.rawMinutes : 0), 0) * 10) / 10,
      totalRoundedMinutes: rows.reduce((sum, row) => sum + (row.valid ? row.roundedMinutes : 0), 0),
      valid: rows.length > 0 && rows.every(row => row.valid)
    };
  }

  static getActiveOperationsForOrder(order, config) {
    const exclusions = (order && order.routeExclusions) || {};
    const source = config || this.readConfig();
    const profile = this.getProfileForFacadeType(order && order.facadeType, source);
    const routed = this.getOperationsForProfile(profile, source);
    return routed.filter(operation =>
      operation.active &&
      !exclusions[operation.daysKey || operation.taskType] &&
      (Boolean(profile && profile.routeOperations && profile.routeOperations.length) ||
        !operation.optionalGroup || (operation.optionalGroup === "lacquer" && Boolean(order && order.lacquerOperations)))
    );
  }

  static getOperationsForProfile(profile, config) {
    const source = config || this.readConfig();
    if (!profile || !profile.routeOperations || !profile.routeOperations.length) return source.operations.slice();
    const byCode = source.operations.reduce((result, operation) => {
      result[operation.taskType] = operation;
      return result;
    }, {});
    return profile.routeOperations.map(code => byCode[code]).filter(Boolean);
  }

  static validateConfiguration(config) {
    const errors = [];
    const activeProfiles = config.profiles.filter(profile => profile.active && profile.facadeTypes.length);
    const activeOperations = config.operations.filter(operation => operation.active);
    const facadeOwners = {};
    activeProfiles.forEach(profile => {
      profile.facadeTypes.forEach(type => {
        if (facadeOwners[type]) errors.push(`Тип фасада «${type}» назначен нескольким профилям.`);
        facadeOwners[type] = profile.code;
      });
    });
    activeProfiles.forEach(profile => {
      const unknownCodes = (profile.routeOperations || []).filter(code => !activeOperations.some(operation => operation.taskType === code));
      unknownCodes.forEach(code => errors.push(`В маршруте профиля «${profile.name}» неизвестная или отключённая операция «${code}».`));
      const profileOperations = this.getOperationsForProfile(profile, config).filter(operation => operation.active);
      if (!profileOperations.length) errors.push(`Для профиля «${profile.name}» не задан маршрут.`);
      profileOperations.forEach(operation => {
        const norm = config.norms.find(item => item.profile === profile.code && item.operation === operation.taskType);
        if (!norm || norm.fixedMinutes === null || norm.minutesPerSqm === null) {
          errors.push(`Заполните норму «${operation.stage}» для профиля «${profile.name}».`);
        }
      });
    });
    if (!activeProfiles.length) errors.push("Добавьте хотя бы один активный профиль с типами фасада.");
    if (!activeOperations.length) errors.push("Добавьте хотя бы одну активную операцию.");
    return { valid: errors.length === 0, errors };
  }

  static saveConfiguration(payload) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const current = this.readConfig();
      const normalized = this.normalizePayload(payload, current);
      if (current.hourlyEnabled) {
        const validation = this.validateConfiguration(normalized);
        if (!validation.valid) throw new Error(validation.errors[0]);
        this.assertActiveOperationsNotDisabled(current, normalized);
        this.assertActiveOrdersCompatible(normalized);
      }
      this.replaceConfigurationRows(normalized, current.hourlyEnabled);
      return this.getClientConfig();
    } finally {
      lock.releaseLock();
    }
  }

  static normalizePayload(payload, current) {
    const currentCodes = new Set(current.operations.map(operation => operation.taskType));
    const profiles = (payload && Array.isArray(payload.profiles) ? payload.profiles : []).map(profile => ({
      code: String(profile.code || "").trim(),
      name: String(profile.name || "").trim(),
      facadeTypes: [...new Set((Array.isArray(profile.facadeTypes) ? profile.facadeTypes : String(profile.facadeTypes || "").split(","))
        .map(value => String(value || "").trim()).filter(Boolean))],
      routeOperations: [...new Set((Array.isArray(profile.routeOperations) ? profile.routeOperations : String(profile.routeOperations || "").split(","))
        .map(value => String(value || "").trim()).filter(Boolean))],
      active: Boolean(profile.active)
    }));
    const operations = (payload && Array.isArray(payload.operations) ? payload.operations : []).map((operation, index) => {
      let code = String(operation.taskType || "").trim();
      if (!code) code = "custom_" + Date.now() + "_" + index;
      if (operation.originalCode && operation.originalCode !== code) throw new Error("Стабильный код операции нельзя изменить.");
      if (currentCodes.has(String(operation.originalCode || code)) && String(operation.originalCode || code) !== code) {
        throw new Error("Стабильный код операции нельзя изменить.");
      }
      const role = operation.role === "sanding" ? "sanding" : "painting";
      return {
        taskType: code,
        stage: String(operation.stage || "").trim(),
        role,
        order: Number(operation.order) || (index + 1) * 10,
        group: operation.group === "lacquer" ? "lacquer" : "main",
        active: Boolean(operation.active),
        daysKey: code,
        column: "",
        exclusionColumn: "",
        optionalGroup: operation.group === "lacquer" ? "lacquer" : ""
      };
    });
    if (profiles.some(profile => !profile.code || !profile.name)) throw new Error("У каждого профиля должны быть код и название.");
    if (operations.some(operation => !operation.taskType || !operation.stage)) throw new Error("У каждой операции должны быть код и название.");
    if (new Set(profiles.map(profile => profile.code)).size !== profiles.length) throw new Error("Коды профилей не должны повторяться.");
    if (new Set(operations.map(operation => operation.taskType)).size !== operations.length) throw new Error("Коды операций не должны повторяться.");
    const norms = [];
    profiles.forEach(profile => operations.forEach(operation => {
      const source = (payload.norms || []).find(item => item.profile === profile.code && item.operation === (operation.taskType || item.operation));
      norms.push({
        profile: profile.code,
        operation: operation.taskType,
        fixedMinutes: this.toNullableNumber(source && source.fixedMinutes),
        minutesPerSqm: this.toNullableNumber(source && source.minutesPerSqm),
        areaBasis: this.normalizeAreaBasis(source && source.areaBasis),
        comment: String(source && source.comment || "").trim()
      });
    }));
    return { schemaVersion: current.schemaVersion, hourlyEnabled: current.hourlyEnabled, profiles, operations, norms };
  }

  static replaceConfigurationRows(config, hourlyEnabled) {
    const sheet = this.ensureSheet();
    const headers = this.getHeaders();
    const build = values => headers.map(header => values[header] === undefined ? "" : values[header]);
    const rows = [
      build({ "Тип записи": "SETTING", "Код": "schema_version", "Значение": config.schemaVersion || "1.3", "Комментарий": "Версия структуры листа." }),
      build({ "Тип записи": "SETTING", "Код": "hourly_planning_enabled", "Значение": hourlyEnabled ? "true" : "false", "Комментарий": "Включается только после успешной проверки норм и активных заказов." })
    ];
    config.profiles.forEach(profile => rows.push(build({
      "Тип записи": "PROFILE", "Код": profile.code, "Название": profile.name,
      "Типы фасада": profile.facadeTypes.join(", "),
      "Операции маршрута": (profile.routeOperations || []).join(", "), "Активна": profile.active
    })));
    config.operations.forEach(operation => rows.push(build({
      "Тип записи": "OPERATION", "Код": operation.taskType, "Название": operation.stage,
      "Роль": operation.role === "sanding" ? "Шлифовка" : "Покраска", "Порядок": operation.order,
      "Группа": operation.group || "main", "Активна": operation.active
    })));
    config.norms.forEach(norm => rows.push(build({
      "Тип записи": "NORM", "Профиль": norm.profile, "Операция": norm.operation,
      "Фиксировано, мин": norm.fixedMinutes === null ? "" : norm.fixedMinutes,
      "На м², мин": norm.minutesPerSqm === null ? "" : norm.minutesPerSqm,
      "База площади": norm.areaBasis || "painting",
      "Комментарий": norm.comment
    })));
    sheet.clearContents();
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    this.formatSheet(sheet);
    this.invalidateCache();
  }

  static assertActiveOperationsNotDisabled(current, next) {
    const disabled = current.operations.filter(operation => operation.active &&
      !next.operations.some(item => item.taskType === operation.taskType && item.active));
    if (!disabled.length) return;
    const activeOrders = this.getActivePaintingOrders();
    const blocked = disabled.find(operation => activeOrders.some(order =>
      String(order.paintingOperationCode || "").trim() === operation.taskType ||
      String(order.stage || "").trim() === operation.stage
    ));
    if (blocked) throw new Error(`Нельзя отключить операцию «${blocked.stage}»: на ней есть активный заказ.`);
  }

  static assertActiveOrdersCompatible(config) {
    this.getActivePaintingOrders().forEach(order => {
      if (!(Number(order.paintingArea) > 0)) throw new Error(`У изделия «${order.displayName}» не указана площадь покраски.`);
      if (!this.getProfileForFacadeType(order.facadeType, config)) {
        throw new Error(`Изделие «${order.displayName}» останется без нормативного профиля.`);
      }
      const operations = this.getActiveOperationsForOrder(order, config);
      const invalid = operations.map(operation => this.calculateOperation(order, operation, config)).find(result => !result.valid);
      if (invalid) throw new Error(`Изделие «${order.displayName}»: ${invalid.message}`);
    });
  }

  static activateHourlyPlanning() {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const config = this.readConfig();
      const validation = this.validateConfiguration(config);
      if (!validation.valid) throw new Error(validation.errors[0]);
      const activeOrders = this.getActivePaintingOrders();
      activeOrders.forEach(order => {
        if (!(Number(order.paintingArea) > 0)) throw new Error(`У изделия «${order.displayName}» не указана площадь покраски.`);
        if (!this.getProfileForFacadeType(order.facadeType, config)) {
          throw new Error(`Для типа фасада «${order.facadeType || "Не выбран"}» изделия «${order.displayName}» не назначен профиль.`);
        }
      });
      this.assertActiveOrdersCompatible(config);
      this.migratePaintingOperationCodes(config);
      this.setSetting("hourly_planning_enabled", "true");
      return this.getClientConfig();
    } finally {
      lock.releaseLock();
    }
  }

  static getActivePaintingOrders() {
    const operations = this.getOperations(false);
    const stageSet = new Set(["В покраску"].concat(operations.map(operation => operation.stage)));
    const codeSet = new Set(operations.map(operation => operation.taskType));
    return SheetService.getObjects(CONFIG.SHEETS.ORDERS)
      .map(row => OrderMapper.map(row))
      .filter(order => (stageSet.has(String(order.stage || "").trim()) || codeSet.has(String(order.paintingOperationCode || "").trim())) &&
        String(order.status || "").trim().toLowerCase() !== "готово");
  }

  static migratePaintingOperationCodes(config) {
    SheetService.ensureOrderColumn("Операция Малярки");
    const sheet = SheetService.getSheet(CONFIG.SHEETS.ORDERS);
    const columns = SheetService.getColumns(sheet);
    if (sheet.getLastRow() < 2 || !columns["Операция Малярки"] || !columns["Этап"]) return;
    const range = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn());
    const values = range.getValues();
    let changed = false;
    values.forEach(row => {
      if (String(row[columns["Операция Малярки"] - 1] || "").trim()) return;
      const stage = String(row[columns["Этап"] - 1] || "").trim();
      const operation = config.operations.find(item => item.stage === stage);
      if (!operation) return;
      row[columns["Операция Малярки"] - 1] = operation.taskType;
      changed = true;
    });
    if (changed) range.setValues(values);
  }

  static setSetting(code, value) {
    const sheet = this.ensureSheet();
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0].map(value => String(value || "").trim());
    const typeIndex = headers.indexOf("Тип записи");
    const codeIndex = headers.indexOf("Код");
    const valueIndex = headers.indexOf("Значение");
    for (let index = 1; index < rows.length; index++) {
      if (String(rows[index][typeIndex] || "").trim() === "SETTING" && String(rows[index][codeIndex] || "").trim() === code) {
        sheet.getRange(index + 1, valueIndex + 1).setValue(value);
        this.invalidateCache();
        return;
      }
    }
    throw new Error("Настройка не найдена: " + code);
  }

  static toBoolean(value) {
    return value === true || ["true", "да", "1"].includes(String(value || "").trim().toLowerCase());
  }

  static toNullableNumber(value) {
    if (value === "" || value === null || value === undefined) return null;
    const number = Number(String(value).replace(",", "."));
    if (!Number.isFinite(number) || number < 0) throw new Error("Норма должна быть числом не меньше нуля.");
    return Math.round(number * 10) / 10;
  }

  static normalizeAreaBasis(value) {
    const basis = String(value || "painting").trim().toLowerCase();
    return ["painting", "back", "front", "polishing"].includes(basis) ? basis : "painting";
  }
}

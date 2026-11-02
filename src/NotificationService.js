const NOTIFICATION_RECIPIENT_HEADERS = ["Сотрудник", "Роль", "Email", "Активен", "Получать уведомления"];
const NOTIFICATION_LOG_HEADERS = [
  "ID события", "Создано", "Ключ изделия", "Договор", "Клиент", "Изделие",
  "Тип события", "Старый этап", "Новый этап", "Старый ответственный",
  "Новый ответственный", "Получатель", "Email получателя", "Статус",
  "Отправлено", "Ошибка", "Тема", "Тело письма", "Ключ идемпотентности"
];

/**
 * Уведомления о новых задачах в конструкторском маршруте.
 * Запускать только после успешного сохранения основной строки заказа.
 */
class NotificationService {
  static initialize() {
    const directory = this.ensureRecipientDirectory();
    const journal = this.ensureLogSheet();
    const settings = this.getSettings();
    return {
      directory: directory.getName(), journal: journal.getName(),
      testMode: settings.testMode, testRecipientConfigured: Boolean(settings.testRecipient)
    };
  }

  static getSettings() {
    const properties = PropertiesService.getScriptProperties();
    return {
      enabled: this.toBoolean(properties.getProperty("EMAIL_NOTIFICATIONS_ENABLED"), true),
      senderName: String(properties.getProperty("EMAIL_SENDER_NAME") || "DCW Производство").trim(),
      testMode: this.toBoolean(properties.getProperty("EMAIL_TEST_MODE"), true),
      testRecipient: String(properties.getProperty("EMAIL_TEST_RECIPIENT") || "").trim(),
      applicationUrl: String(properties.getProperty("DCW_APPLICATION_URL") || this.getApplicationUrl()).trim()
    };
  }

  static toBoolean(value, fallback) {
    if (value === null || value === undefined || value === "") return fallback;
    return ["1", "true", "yes", "да"].includes(String(value).trim().toLowerCase());
  }

  static getApplicationUrl() {
    try { return ScriptApp.getService().getUrl() || ""; }
    catch (error) { return ""; }
  }

  static onAssignmentChanged(oldItem, newItem, context) {
    try {
      const definition = this.getEventDefinition(context && context.eventType);
      if (!definition) return { status: "SKIPPED", reason: "Событие не требует уведомления." };

      const eventId = Utilities.getUuid();
      const item = Object.assign({}, newItem || {});
      const idempotencyKey = this.buildIdempotencyKey(oldItem, item, definition);
      if (this.wasLastEventDuplicate(item.key, idempotencyKey)) {
        return { eventId, status: "SKIPPED", reason: "Повторное выполнение того же перехода." };
      }
      const resolution = this.resolveRecipients(item, definition.recipientType);
      if (!resolution.recipients.length) {
        const notification = this.buildNotification(item, definition, "", "");
        return this.recordAndSend(eventId, oldItem, item, definition, notification,
          "SKIPPED", "Не найден активный получатель с включёнными уведомлениями. " + resolution.diagnostic, idempotencyKey);
      }

      return resolution.recipients.map(recipient => {
        const notification = this.buildNotification(item, definition, recipient.name, recipient.email);
        return this.recordAndSend(eventId, oldItem, item, definition, notification, "PENDING", "", idempotencyKey);
      });
    } catch (error) {
      Logger.log("NOTIFICATION_ERROR " + (error && error.stack || error));
      return { status: "FAILED", error: String(error && error.message || error) };
    }
  }

  static getEventDefinition(eventType) {
    return {
      TECHNOLOGIST_REVIEW: { code: "TECHNOLOGIST_REVIEW", recipientType: "technologist", action: "Проверить ТЗ" },
      TECHNOLOGIST_REVISION: { code: "TECHNOLOGIST_REVISION", recipientType: "manager", action: "Доработать ТЗ и повторно передать технологу" },
      PROFORMA_ASSIGNED: { code: "PROFORMA_ASSIGNED", recipientType: "constructor", action: "Подготовить проформу" },
      PROFORMA_APPROVAL: { code: "PROFORMA_APPROVAL", recipientType: "manager", action: "Согласовать проформу с клиентом" },
      KD_ASSIGNED: { code: "KD_ASSIGNED", recipientType: "constructor", action: "Разработать КД" },
      CONSTRUCTOR_REASSIGNED: { code: "CONSTRUCTOR_REASSIGNED", recipientType: "constructor", action: "Продолжить работу по текущей задаче" }
    }[String(eventType || "").trim()] || null;
  }

  static resolveRecipients(item, recipientType) {
    const employees = this.getEmployees();
    const expectedRoles = {
      manager: ["менедж"], constructor: ["конструкт"], technologist: ["технолог", "техн"]
    }[recipientType] || [];
    let candidates;
    if (recipientType === "technologist") {
      candidates = employees.filter(employee => this.matchesRole(employee.role, expectedRoles));
    } else {
      const name = String(recipientType === "manager" ? item.manager : item.constructor || "").trim();
      candidates = employees.filter(employee =>
        this.normalizeText(employee.name) === this.normalizeText(name) && this.matchesRole(employee.role, expectedRoles)
      );
    }
    const unique = {};
    candidates.forEach(employee => { unique[employee.email.toLowerCase()] = employee; });
    return {
      recipients: Object.keys(unique).map(key => unique[key]),
      diagnostic: "Тип: " + recipientType + ". Активные записи: " +
        (employees.map(employee => employee.name + " (" + employee.role + ")").join(", ") || "нет") + "."
    };
  }

  static matchesRole(role, expectedRoles) {
    const normalized = this.normalizeText(role);
    return expectedRoles.some(expected => normalized.includes(this.normalizeText(expected)));
  }

  static normalizeText(value) {
    return String(value == null ? "" : value).replace(/\u00a0/g, " ").trim().replace(/\s+/g, " ").toLowerCase();
  }

  static getEmployees() {
    const sheet = this.ensureRecipientDirectory();
    const rows = sheet.getDataRange().getValues();
    const columns = {};
    (rows[0] || []).forEach((header, index) => { columns[String(header || "").trim()] = index; });
    return rows.slice(1).map(row => ({
      name: String(row[columns["Сотрудник"]] || "").trim(),
      role: String(row[columns["Роль"]] || "").trim(),
      email: String(row[columns["Email"]] || "").trim(),
      active: this.toBoolean(row[columns["Активен"]], false),
      enabled: this.toBoolean(row[columns["Получать уведомления"]], false)
    })).filter(employee => employee.name && employee.role && employee.email && employee.active && employee.enabled);
  }

  static getRecipientDiagnostics() {
    const employees = this.getEmployees();
    return {
      activeEmployees: employees.map(employee => ({ name: employee.name, role: employee.role })),
      technologists: this.resolveRecipients({}, "technologist").recipients.map(employee => employee.name)
    };
  }

  static ensureRecipientDirectory() {
    const sheet = SheetService.getSheet(CONFIG.SHEETS.DICTIONARIES);
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0]
      .map(value => String(value || "").trim());
    const missing = NOTIFICATION_RECIPIENT_HEADERS.filter(header => !headers.includes(header));
    if (missing.length) sheet.getRange(1, sheet.getLastColumn() + 1, 1, missing.length).setValues([missing]);
    return sheet;
  }

  static ensureLogSheet() {
    const spreadsheet = SheetService.getSpreadsheet();
    let sheet = spreadsheet.getSheetByName(CONFIG.SHEETS.NOTIFICATION_LOG);
    if (!sheet) {
      sheet = spreadsheet.insertSheet(CONFIG.SHEETS.NOTIFICATION_LOG);
      sheet.getRange(1, 1, 1, NOTIFICATION_LOG_HEADERS.length).setValues([NOTIFICATION_LOG_HEADERS]);
      sheet.getRange(1, 1, 1, NOTIFICATION_LOG_HEADERS.length).setFontWeight("bold").setBackground("#e2e8f0");
      sheet.setFrozenRows(1);
    }
    return sheet;
  }

  static buildNotification(item, definition, recipientName, recipientEmail) {
    const settings = this.getSettings();
    const lines = [
      ["Договор", item.contract], ["Клиент", item.client], ["Изделие", item.product],
      ["Этап", item.stage], ["Ответственный", this.getResponsibleLabel(item, definition.recipientType)],
      ["Требуемое действие", definition.action], ["Плановая дата завершения", item.plannedDate || item.deadline]
    ].filter(pair => String(pair[1] || "").trim());
    const comment = String(item.comment || "").trim();
    const appLink = settings.applicationUrl;
    const subject = ["DCW", definition.action || "Новое действие", item.client || "", item.product || "",
      item.contract ? "Договор №" + item.contract : ""]
      .filter(Boolean).join(" | ");
    const text = [
      "Здравствуйте!", "", "Вам передано изделие, требующее выполнения следующего действия.", "",
      ...lines.map(pair => pair[0] + ": " + pair[1]),
      ...(comment ? ["", "Комментарий:", comment] : []),
      ...(appLink ? ["", "Открыть изделие: " + appLink] : []),
      "", "---", "Автоматическое уведомление системы DCW."
    ].join("\n");
    const html = [
      "Здравствуйте!", "", "Вам передано изделие, требующее выполнения следующего действия.", "",
      ...lines.map(pair => "<b>" + this.escapeHtml(pair[0]) + ":</b> " + this.escapeHtml(pair[1])),
      ...(comment ? ["", "<b>Комментарий:</b>", this.escapeHtml(comment)] : []),
      ...(appLink ? ["", '<a href="' + this.escapeHtml(appLink) + '">Открыть изделие</a>'] : []),
      "", "---", "Автоматическое уведомление системы DCW."
    ].join("<br>");
    return { recipientName, recipientEmail, subject, text, html, senderName: settings.senderName };
  }

  static buildIdempotencyKey(oldItem, newItem, definition) {
    return [newItem && newItem.key, definition.code, oldItem && oldItem.stage, newItem && newItem.stage,
      oldItem && oldItem.responsible, newItem && newItem.responsible,
      oldItem && oldItem.constructor, newItem && newItem.constructor]
      .map(value => String(value || "").trim()).join("|");
  }

  static wasLastEventDuplicate(orderKey, idempotencyKey) {
    if (!orderKey || !idempotencyKey) return false;
    const sheet = this.ensureLogSheet();
    const values = sheet.getDataRange().getDisplayValues();
    const headers = values[0] || [];
    const keyColumn = headers.indexOf("Ключ изделия");
    const idempotencyColumn = headers.indexOf("Ключ идемпотентности");
    if (keyColumn < 0 || idempotencyColumn < 0) return false;
    for (let index = values.length - 1; index > 0; index -= 1) {
      if (String(values[index][keyColumn] || "") !== String(orderKey)) continue;
      return String(values[index][idempotencyColumn] || "") === idempotencyKey;
    }
    return false;
  }

  static getResponsibleLabel(item, recipientType) {
    if (recipientType === "manager") return item.manager;
    if (recipientType === "constructor") return item.constructor;
    return "Технолог";
  }

  static recordAndSend(eventId, oldItem, newItem, definition, notification, initialStatus, initialError, idempotencyKey) {
    const sheet = this.ensureLogSheet();
    const rowNumber = sheet.getLastRow() + 1;
    sheet.getRange(rowNumber, 1, 1, NOTIFICATION_LOG_HEADERS.length).setValues([[
      eventId, new Date(), newItem.key || "", newItem.contract || "", newItem.client || "", newItem.product || "",
      definition.code, oldItem && oldItem.stage || "", newItem.stage || "",
      oldItem && oldItem.responsible || "", newItem.responsible || "", notification.recipientName || "",
      notification.recipientEmail || "", initialStatus, "", initialError || "", notification.subject || "",
      notification.text || "", idempotencyKey || ""
    ]]);
    if (initialStatus !== "PENDING") return { eventId, status: initialStatus, reason: initialError };

    const result = this.sendNotification(notification);
    this.updateLogRow(sheet, rowNumber, result.status, result.error || "", result.status === "SENT" ? new Date() : "");
    return { eventId, status: result.status, error: result.error || "" };
  }

  static sendNotification(notification) {
    const settings = this.getSettings();
    if (!settings.enabled) return { status: "SKIPPED", error: "Уведомления отключены настройкой." };
    const recipient = settings.testMode ? settings.testRecipient : notification.recipientEmail;
    if (!recipient) {
      return { status: "SKIPPED", error: settings.testMode
        ? "Тестовый адрес EMAIL_TEST_RECIPIENT не задан."
        : "У получателя отсутствует email." };
    }
    try {
      MailApp.sendEmail({ to: recipient, subject: notification.subject, body: notification.text,
        htmlBody: notification.html, name: settings.senderName });
      return { status: "SENT" };
    } catch (error) {
      return { status: "FAILED", error: String(error && error.message || error).slice(0, 1000) };
    }
  }

  // Не отправляет письмо: вызов MailApp нужен только для выдачи OAuth-разрешения.
  static requestMailAuthorization() {
    return { remainingDailyQuota: MailApp.getRemainingDailyQuota() };
  }

  static updateLogRow(sheet, rowNumber, status, error, sentAt) {
    const columns = { status: NOTIFICATION_LOG_HEADERS.indexOf("Статус") + 1, sentAt: NOTIFICATION_LOG_HEADERS.indexOf("Отправлено") + 1,
      error: NOTIFICATION_LOG_HEADERS.indexOf("Ошибка") + 1 };
    sheet.getRange(rowNumber, columns.status, 1, 3).setValues([[status, sentAt, error]]);
  }

  static retryFailedNotifications(limit) {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
      const sheet = this.ensureLogSheet();
      const values = sheet.getDataRange().getValues();
      const headers = values[0] || [];
      const column = name => headers.indexOf(name);
      const max = Math.max(1, Math.min(Number(limit) || 20, 100));
      let retried = 0;
      values.slice(1).forEach((row, index) => {
        if (retried >= max || String(row[column("Статус")] || "") !== "FAILED") return;
        retried += 1;
        const notification = { recipientEmail: String(row[column("Email получателя")] || ""),
          subject: String(row[column("Тема")] || ""), text: String(row[column("Тело письма")] || ""),
          html: this.escapeHtml(String(row[column("Тело письма")] || "")).replace(/\n/g, "<br>"),
          senderName: this.getSettings().senderName };
        const result = this.sendNotification(notification);
        this.updateLogRow(sheet, index + 2, result.status, result.error || "", result.status === "SENT" ? new Date() : "");
      });
      return { retried };
    } finally {
      lock.releaseLock();
    }
  }

  static escapeHtml(value) {
    return String(value == null ? "" : value).replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
  }
}

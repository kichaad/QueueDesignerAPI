/**
 * Работа с Облаком Mail.ru через Cloudflare Worker.
 * Секрет Mail.ru хранится в Worker, а не в Apps Script.
 */
const MAIL_CLOUD_API_URL_PROPERTY = "mail.cloud.api.url";
const MAIL_CLOUD_API_TOKEN_PROPERTY = "mail.cloud.api.token";
const MAIL_CLOUD_ROOT_PATH = "/DCW/В РАБОТЕ/Пректы для ИИ";

class MailCloudService {

  static createOrderFolder(order) {
    const folderName = this.getOrderFolderName(order);
    const path = MAIL_CLOUD_ROOT_PATH + "/" + folderName;
    this.callWorker("create-folder", order);

    return {
      name: folderName,
      url: "https://cloud.mail.ru/home" + this.encodePath(path)
    };
  }

  static getOrderFolderName(order) {
    // Имя строится из составного ключа строки заказа. Символ «|» заменяем,
    // так как он недопустим в именах папок на части клиентов WebDAV.
    return this.getSafeName(order.key || [
      order.contract,
      order.client,
      order.product
    ].join("|"));
  }

  static getOrderPath(order) {
    return MAIL_CLOUD_ROOT_PATH + "/" + this.getOrderFolderName(order);
  }

  static getOrderFiles(order) {
    return this.callWorker("list", order).files || [];
  }

  static uploadOrderFile(order, file) {
    return this.callWorker("upload", order, file);
  }

  static callWorker(action, order, file) {
    const properties = PropertiesService.getScriptProperties();
    const url = properties.getProperty(MAIL_CLOUD_API_URL_PROPERTY);
    const token = properties.getProperty(MAIL_CLOUD_API_TOKEN_PROPERTY);
    if (!url || !token) {
      throw new Error("Не настроен Worker Mail.ru. Добавьте свойства «mail.cloud.api.url» и «mail.cloud.api.token» в Apps Script.");
    }
    const response = UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      headers: { Authorization: "Bearer " + token },
      payload: JSON.stringify({ action: action, orderKey: order.key, file: file || null }),
      muteHttpExceptions: true
    });
    const data = JSON.parse(response.getContentText() || "{}");
    if (response.getResponseCode() < 200 || response.getResponseCode() >= 300) {
      throw new Error(data.error || "Ошибка Worker Mail.ru (код " + response.getResponseCode() + ").");
    }
    return data;
  }

  static encodePath(path) {
    return String(path).split("/").map(part =>
      part ? encodeURIComponent(part) : ""
    ).join("/");
  }

  static getSafeName(value) {
    return String(value || "Заказ")
      .replace(/[\\/:*?"<>|]/g, "_")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 180) || "Заказ";
  }
}

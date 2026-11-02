/**
 * Хранилище файлов заказов в Google Drive.
 * Каждому заказу соответствует отдельная папка с именем из его ключа.
 */
const GOOGLE_DRIVE_ORDERS_ROOT_ID = "1D0OCPZoA6pXFBUPGDKUrBEJCQcZOXvk_";
const GOOGLE_DRIVE_FILE_SECTIONS = ["ТЗ", "Замер", "Проформа", "КД"];

class GoogleDriveService {

  static createOrderFolder(order) {
    const folderName = this.getOrderFolderName(order);
    const root = DriveApp.getFolderById(GOOGLE_DRIVE_ORDERS_ROOT_ID);
    const matches = root.getFoldersByName(folderName);
    const folder = matches.hasNext() ? matches.next() : root.createFolder(folderName);
    return this.toFolderInfo(folder);
  }

  static getOrderFiles(order) {
    const folders = this.ensureOrderFileFolders(order);
    return GOOGLE_DRIVE_FILE_SECTIONS.reduce((result, section) => {
      result[section] = this.listFolderFiles(folders[section]);
      return result;
    }, {});
  }

  static uploadOrderFile(order, section, file) {
    if (!file || !file.name || !file.base64) {
      throw new Error("Не удалось получить файл для загрузки.");
    }

    const folder = this.ensureOrderFileFolders(order)[this.normalizeSection(section)];
    const bytes = Utilities.base64Decode(String(file.base64));
    if (bytes.length > 20 * 1024 * 1024) throw new Error("Файл больше 20 МБ.");
    const blob = Utilities.newBlob(
      bytes,
      file.mimeType || "application/octet-stream",
      this.getSafeFileName(file.name)
    );
    const created = folder.createFile(blob);
    return {
      id: created.getId(),
      name: created.getName(),
      url: created.getUrl(),
      mimeType: created.getMimeType()
    };
  }

  static downloadOrderFile(order, section, fileId) {
    const folder = this.ensureOrderFileFolders(order)[this.normalizeSection(section)];
    const file = this.getFileFromFolder(folder, fileId);
    const maxBytes = 20 * 1024 * 1024;
    if (file.getSize() > maxBytes) {
      throw new Error("Для скачивания через приложение доступен файл до 20 МБ.");
    }
    const blob = file.getBlob();
    return {
      name: file.getName(),
      mimeType: blob.getContentType() || "application/octet-stream",
      base64: Utilities.base64Encode(blob.getBytes())
    };
  }

  static downloadMeasurementArchive(order) {
    const folder = this.ensureOrderFileFolders(order)["Замер"];
    const files = [];
    const iterator = folder.getFiles();
    let total = 0;
    const maxBytes = 10 * 1024 * 1024;
    while (iterator.hasNext()) {
      const file = iterator.next();
      total += file.getSize();
      if (total > maxBytes) throw new Error("Суммарный размер файлов замера больше 10 МБ. Скачайте файлы по отдельности.");
      files.push(file);
    }
    if (!files.length) throw new Error("В разделе «Замер» нет файлов для скачивания.");
    files.sort((a, b) => a.getName().localeCompare(b.getName(), "ru") || a.getId().localeCompare(b.getId()));
    const used = new Set();
    let blobTotal = 0;
    const blobs = files.map(file => {
      const name = this.getSafeFileName(file.getName());
      const dot = name.lastIndexOf(".");
      const base = dot > 0 ? name.slice(0, dot) : name;
      const ext = dot > 0 ? name.slice(dot) : "";
      let candidate = name;
      let number = 2;
      while (used.has(candidate.toLowerCase())) candidate = base + " (" + number++ + ")" + ext;
      used.add(candidate.toLowerCase());
      const blob = file.getBlob();
      blobTotal += blob.getBytes().length;
      if (blobTotal > maxBytes) throw new Error("Суммарный размер файлов замера больше 10 МБ. Скачайте файлы по отдельности.");
      return blob.setName(candidate);
    });
    const zip = Utilities.zip(blobs, "Замер.zip");
    if (zip.getBytes().length > maxBytes) throw new Error("Архив больше 10 МБ. Скачайте файлы по отдельности.");
    return { name: "Замер.zip", mimeType: "application/zip", base64: Utilities.base64Encode(zip.getBytes()) };
  }

  static getOrderFileThumbnail(order, section, fileId) {
    const folder = this.ensureOrderFileFolders(order)[this.normalizeSection(section)];
    const file = this.getFileFromFolder(folder, fileId);
    const mimeType = String(file.getMimeType() || "").toLowerCase();
    if (!mimeType.startsWith("image/")) {
      throw new Error("Миниатюра доступна только для изображений.");
    }
    const thumbnail = file.getThumbnail() || file.getBlob();
    return {
      mimeType: thumbnail.getContentType() || "image/jpeg",
      base64: Utilities.base64Encode(thumbnail.getBytes())
    };
  }

  static deleteOrderFile(order, section, fileId) {
    const folder = this.ensureOrderFileFolders(order)[this.normalizeSection(section)];
    const file = this.getFileFromFolder(folder, fileId);
    const result = {
      id: file.getId(),
      name: file.getName()
    };
    file.setTrashed(true);
    return result;
  }

  static ensureOrderFolder(order) {
    const storedId = this.extractFolderId(order.cloudFolderUrl);
    if (storedId) {
      try {
        return DriveApp.getFolderById(storedId);
      } catch (error) {
        Logger.log("Папка Google Drive недоступна, будет создана новая: " + error.message);
      }
    }

    const info = this.createOrderFolder(order);
    SheetService.ensureOrderColumn("Папка Google Drive");
    SheetService.updateField(order.key, "Папка Google Drive", info.url);
    order.cloudFolderUrl = info.url;
    return DriveApp.getFolderById(info.id);
  }

  /**
   * Создаёт стандартную структуру папок проекта. Ранее загруженные файлы
   * лежали в корне проекта, поэтому при первом обращении переносим их в «ТЗ».
   */
  static ensureOrderFileFolders(order) {
    const orderFolder = this.ensureOrderFolder(order);
    return this.ensureProjectFileFolders(orderFolder);
  }

  static ensureProjectFileFolders(orderFolder) {
    const folders = GOOGLE_DRIVE_FILE_SECTIONS.reduce((result, section) => {
      const matches = orderFolder.getFoldersByName(section);
      result[section] = matches.hasNext() ? matches.next() : orderFolder.createFolder(section);
      return result;
    }, {});

    const legacyFiles = [];
    const iterator = orderFolder.getFiles();
    while (iterator.hasNext()) legacyFiles.push(iterator.next());
    legacyFiles.forEach(file => file.moveTo(folders["ТЗ"]));

    return folders;
  }

  // Однократная миграция ранее созданных папок проектов. Безопасно запускать
  // повторно: после первого запуска в корнях проектов уже не останется файлов.
  static migrateExistingOrderFolders() {
    const root = DriveApp.getFolderById(GOOGLE_DRIVE_ORDERS_ROOT_ID);
    const iterator = root.getFolders();
    let folders = 0;
    let movedFiles = 0;
    while (iterator.hasNext()) {
      const projectFolder = iterator.next();
      const before = projectFolder.getFiles();
      while (before.hasNext()) {
        before.next();
        movedFiles += 1;
      }
      this.ensureProjectFileFolders(projectFolder);
      folders += 1;
    }
    return { folders: folders, movedFiles: movedFiles };
  }

  static normalizeSection(section) {
    const value = String(section || "").trim();
    if (!GOOGLE_DRIVE_FILE_SECTIONS.includes(value)) {
      throw new Error("Выберите раздел файлов: ТЗ, Замер, Проформа или КД.");
    }
    return value;
  }

  static listFolderFiles(folder) {
    const files = [];
    const iterator = folder.getFiles();
    while (iterator.hasNext()) {
      const file = iterator.next();
      files.push({
        id: file.getId(),
        name: file.getName(),
        url: file.getUrl(),
        mimeType: file.getMimeType()
      });
    }
    return files.sort((left, right) => left.name.localeCompare(right.name, "ru"));
  }

  static getOrderFolderName(order) {
    // В отличие от старого WebDAV-хранилища Mail.ru, Google Drive поддерживает
    // разделитель «|», поэтому имя папки совпадает с ключом строки заказа.
    return this.getSafeFolderName(order.key || [
      order.contract,
      order.client,
      order.product
    ].join("|"));
  }

  static toFolderInfo(folder) {
    return {
      id: folder.getId(),
      name: folder.getName(),
      url: folder.getUrl()
    };
  }

  static extractFolderId(url) {
    const match = String(url || "").match(/folders\/([\w-]+)/);
    return match ? match[1] : "";
  }

  static getFileFromFolder(folder, fileId) {
    const iterator = folder.getFiles();
    while (iterator.hasNext()) {
      const file = iterator.next();
      if (file.getId() === String(fileId || "")) return file;
    }
    throw new Error("Файл не найден в папке заказа.");
  }

  static getSafeFileName(value) {
    return String(value || "Заказ")
      .replace(/[\\/\u0000]/g, "_")
      .trim()
      .slice(0, 180) || "Заказ";
  }

  static getSafeFolderName(value) {
    return String(value || "Заказ")
      .replace(/[\\/\u0000]/g, "_")
      .replace(/\s+/g, " ")
      .trim()
      .slice(0, 180) || "Заказ";
  }
}

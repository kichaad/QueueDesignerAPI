function doGet(e) {
  // Если параметр resource существует
  if (e && e.parameter && e.parameter.resource) {
    return Api.handleGet(e);
  }

  // Иначе открываем приложение
  const template = HtmlService.createTemplateFromFile('Index');
  template.appUrl = ScriptApp.getService().getUrl();
  return template
    .evaluate()
    .setTitle('Очередь конструкторов')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  return Api.handlePost(e);
}

function addOrder(orderData) {
  return OrderService.createOrder(orderData);
}

function moveOrderToDate(key, targetDate) {
  return OrderService.moveOrderToDate(key, targetDate);
}

function clearManualOrderDate(key) {
  return OrderService.clearManualOrderDate(key);
}

function assignOrderConstructor(key, constructor) {
  return OrderService.assignOrderConstructor(key, constructor);
}

function changeOrderDuration(key, duration) {
  return OrderService.changeOrderDuration(key, duration);
}

function advanceManagerStage(key, targetStage) {
  return OrderService.advanceManagerStage(key, targetStage);
}

function reviewByTechnologist(key, action, constructor, proformaDays, kdDays) {
  return OrderService.reviewByTechnologist(key, action, constructor, proformaDays, kdDays);
}

function changeWorkStatus(key, action, comment) {
  return OrderService.changeWorkStatus(key, action, comment);
}

function changeConstructorTaskDays(key, taskType, days) {
  return OrderService.changeConstructorTaskDays(key, taskType, days);
}

function changeProductionStatus(key, action) {
  return OrderService.changeProductionStatus(key, action);
}

function sendPaintingBranchToPackaging(key, branch) {
  return OrderService.sendPaintingBranchToPackaging(key, branch);
}

function launchManufacturing(key) {
  return OrderService.launchManufacturing(key);
}

function changeManufacturingStatus(key, branch, action, comment) {
  return OrderService.changeManufacturingStatus(key, branch, action, comment);
}

function sendSgpToInstallation(key, branch) {
  return OrderService.sendSgpToInstallation(key, branch);
}

function saveInstallationComment(key, branch, comment) {
  return OrderService.saveInstallationComment(key, branch, comment);
}

function saveClaimComment(key, comment) {
  return OrderService.saveClaimComment(key, comment);
}

function changeManufacturingStage(key, branch, stage) {
  return OrderService.changeManufacturingStage(key, branch, stage);
}

function migrateDeliveryStagesToSgp() {
  return OrderService.ensureSgpStageMigration();
}

function moveOrderToOverallStage(key, stage) {
  return OrderService.moveOrderToOverallStage(key, stage);
}

function sendOrderToClaims(key, branch, comment) {
  return OrderService.sendOrderToClaims(key, branch, comment);
}

function completeClaim(key) {
  return OrderService.completeClaim(key);
}

function setMdfFacade(key, enabled) {
  return OrderService.setMdfFacade(key, enabled);
}

function setVeneerFacade(key, enabled) {
  return OrderService.setVeneerFacade(key, enabled);
}

function setFacadeMaterials(key, selection) {
  return OrderService.setFacadeMaterials(key, selection);
}

function setCorpusTypes(key, selection) {
  return OrderService.setCorpusTypes(key, selection);
}

function setMaterialAreas(key, areas) {
  return OrderService.setMaterialAreas(key, areas);
}

function saveProductionDays(key, days) {
  return OrderService.saveProductionDays(key, days);
}

function getPaintingNorms() {
  return PaintingNormService.getClientConfig();
}

function savePaintingNorms(payload) {
  return PaintingNormService.saveConfiguration(payload);
}

function activateHourlyPaintingPlanning() {
  return PaintingNormService.activateHourlyPlanning();
}

function changeFacadeType(key, facadeType) {
  return OrderService.changeFacadeType(key, facadeType);
}

function setProformaApprovedDate(key, value) {
  return OrderService.setProformaApprovedDate(key, value);
}

function setOrderDeadlineSettings(key, mode, manufacturingDays, manualDate) {
  return OrderService.setOrderDeadlineSettings(key, mode, manufacturingDays, manualDate);
}

function changeSandingStatus(key, taskType, action, comment, branch) {
  return OrderService.changeSandingStatus(key, taskType, action, comment, branch);
}

function changeSandingWorker(key, worker) {
  return OrderService.changeSandingWorker(key, worker);
}

function setAvailableSandingWorkers(workers) {
  return OrderService.setAvailableSandingWorkers(workers);
}

function changePaintingStatus(key, taskType, action, comment, branch) {
  return OrderService.changePaintingStatus(key, taskType, action, comment, branch);
}

function repairCorpusPaintingLeak(key) {
  return OrderService.repairCorpusPaintingLeak(key);
}

function migratePendingMdfPaintingBranches() {
  return OrderService.migratePendingMdfPaintingBranches();
}

function migrateLegacyFacadePackagingToMdfBuffer() {
  return OrderService.migrateLegacyFacadePackagingToMdfBuffer();
}

function migrateOverallStages(dryRun) {
  return OrderService.migrateOverallStages(dryRun !== false);
}

function migrateMissingVeneerFacadeToPaintingBuffer(dryRun) {
  return OrderService.migrateMissingVeneerFacadeToPaintingBuffer(dryRun !== false);
}

function migrateClaimedInstallationBranchesToFinished() {
  return OrderService.migrateClaimedInstallationBranchesToFinished();
}

function getOrderFiles(key) {
  return OrderService.getOrderFiles(key);
}

function uploadOrderFile(key, section, file) {
  return OrderService.uploadOrderFile(key, section, file);
}

function downloadOrderFile(key, section, fileId) {
  return OrderService.downloadOrderFile(key, section, fileId);
}

function downloadMeasurementArchive(key) {
  return OrderService.downloadMeasurementArchive(key);
}

function getOrderFileThumbnail(key, section, fileId) {
  return OrderService.getOrderFileThumbnail(key, section, fileId);
}

function deleteOrderFile(key, section, fileId) {
  return OrderService.deleteOrderFile(key, section, fileId);
}

function getMailCloudConnectionStatus() {
  return {
    scriptId: ScriptApp.getScriptId(),
    passwordConfigured: Boolean(PropertiesService.getScriptProperties().getProperty("cloud.mail.ru"))
  };
}

// Однократная служебная команда для переноса прежних файлов из корня каждой
// папки проекта в подпапку «ТЗ» и создания структуры ТЗ / Проформа / КД.
function migrateExistingOrderFileFolders() {
  return GoogleDriveService.migrateExistingOrderFolders();
}

// Однократно запустите эту функцию из редактора Apps Script после публикации,
// чтобы Google запросил у владельца приложения доступ к папке заказов.
function authorizeGoogleDriveStorage() {
  const folder = DriveApp.getFolderById(GOOGLE_DRIVE_ORDERS_ROOT_ID);
  return { name: folder.getName(), url: folder.getUrl() };
}

function testOrders() {
  const orders = OrderService.getOrders("constructors");
  Logger.log(JSON.stringify(orders, null, 2));
}

/**
 * Возвращает список изделий для интерфейса
 */
function assertClientAppVersion(clientVersion) {
  const client = String(clientVersion || '').trim();
  const server = String(CONFIG.APP_VERSION || '').trim();
  if (client !== server) {
    throw new Error('APP_VERSION_MISMATCH|client=' + (client || 'unknown') + '|server=' + server);
  }
}

function getOrders(section, clientVersion) {
  assertClientAppVersion(clientVersion);
  const snapshot = OrderService.getOrders(section);
  snapshot.appVersion = CONFIG.APP_VERSION;
  return snapshot;
}

function getSupplyImportCandidates(fileName) {
  return SupplyService.getImportOrderCandidates(fileName);
}

function importSupplySpecification(file, orderKey) {
  const result = SupplyService.importSpecification(file, orderKey);
  return { result: result, snapshot: SupplyService.getSnapshot() };
}

function previewSupplySpecification(file, orderKey, selections) {
  return SupplyService.previewSpecification(file, orderKey, selections);
}

function analyzeSupplySpecification(file, orderKey) {
  return SupplyService.analyzeSpecification(file, orderKey);
}

function saveSupplySpecification(materials, fileName, orderKey, importInfo) {
  const result = SupplyService.saveSpecification(materials, fileName, orderKey, importInfo);
  return { result: result, snapshot: SupplyService.getSnapshot() };
}

function changePurchaseStatus(rowId, direction) {
  const result = SupplyService.changeStatus(rowId, direction);
  return { result: result, snapshot: SupplyService.getSnapshot() };
}

function setPurchaseStatus(rowId, status) {
  const result = SupplyService.setStatus(rowId, status);
  return { result: result, snapshot: SupplyService.getSnapshot() };
}

function setSupplyArticle(rowId, article, expectedArticle) {
  const result = SupplyService.setArticle(rowId, article, expectedArticle);
  return { result: result, snapshot: SupplyService.getSnapshot() };
}

function assignSupplySupplier(rowId, supplier) {
  const result = SupplyService.assignSupplier(rowId, supplier);
  return { result: result, snapshot: SupplyService.getSnapshot() };
}

function deleteSupplyMaterial(rowId) {
  const result = SupplyService.deleteMaterial(rowId);
  return { result: result, snapshot: SupplyService.getSnapshot() };
}

function getSupplyOrderDirectory() {
  return SupplyService.getMaterialOrderDirectory();
}

function getManualRequestOrderDirectory() { return SupplyService.getManualRequestOrderDirectory(); }
function getManualRequestChoices() { return { orders: SupplyService.getManualRequestOrderDirectory(), clients: EquipmentService.getClients() }; }

function saveManualSupplyRequest(data) {
  const result = SupplyService.saveManualRequest(data);
  return { result: result, snapshot: SupplyService.getSnapshot() };
}

function clearSupplyOrderMaterials(orderKey) {
  const result = SupplyService.clearOrderMaterials(orderKey);
  return { result: result, snapshot: SupplyService.getSnapshot() };
}

function resetMaterialMappingsToManualOnly() {
  return SupplyService.resetMaterialMappingsToManualOnly();
}

function getSupplierOrderCandidates(supplier, productKeys) {
  return SupplierOrderService.getCandidates(supplier, productKeys);
}

function createSupplierOrder(supplier, rowIds) {
  const result = SupplierOrderService.createFromSupplier(supplier, rowIds);
  return { result: result, snapshot: SupplyService.getSnapshot() };
}

function getSupplierOrderDetails(orderId) {
  return SupplierOrderService.getOrderDetails(orderId);
}
function saveSupplierOrderDeliveryDates(orderId, payload, clientVersion) {
  assertClientAppVersion(clientVersion);
  const result = SupplierOrderService.saveDeliveryDates(orderId, payload);
  return { result: result, snapshot: SupplierOrderService.getSnapshot(), appVersion: CONFIG.APP_VERSION };
}
function setSupplierOrderLineArticle(orderId, lineIds, article, expectedArticle, clientVersion) {
  assertClientAppVersion(clientVersion);
  const result = SupplierOrderService.setLineArticle(orderId, lineIds, article, expectedArticle);
  return { result: result, snapshot: SupplierOrderService.getSnapshot(), appVersion: CONFIG.APP_VERSION };
}
function setSupplierOrderLineMoney(orderId, lineIds, field, value, expected, clientVersion) {
  assertClientAppVersion(clientVersion);
  const result = SupplierOrderService.setLineMoney(orderId, lineIds, field, value, expected);
  return { result: result, snapshot: SupplierOrderService.getSnapshot(), appVersion: CONFIG.APP_VERSION };
}
function prepareMdmSupplierOrdersExport(orderIds, clientVersion) {
  assertClientAppVersion(clientVersion);
  return SupplierOrderService.prepareMdmExport(orderIds);
}
function completeMdmSupplierOrdersExport(token, clientVersion) {
  assertClientAppVersion(clientVersion);
  const result = SupplierOrderService.completeMdmExport(token);
  return { result: result, snapshot: SupplierOrderService.getSnapshot(), appVersion: CONFIG.APP_VERSION };
}
function exportSupplierOrderExcel(orderId) {
  return SupplierOrderService.exportExcel(orderId);
}
function exportAllOrdersExcel(rows, filters) {
  return OrderService.exportAllOrdersExcel(rows, filters);
}
function exportProductionExcel(rows, filters, layout) {
  return OrderService.exportProductionExcel(rows, filters, layout);
}
function createBasisSpecification(file) {
  return BasisSpecificationService.createExcel(file);
}
function exportListExcel(exportType, rows, filters) {
  return OrderService.exportListExcel(exportType, rows, filters);
}
function getSupplierOrderAddCandidates(orderId) {
  const order = SupplierOrderService.getOrders().find(item => item.id === String(orderId || "").trim());
  if (!order) throw new Error("Заказ поставщику не найден.");
  if (order.status !== "Сформирован") throw new Error("Состав можно менять только у сформированного заказа.");
  return SupplierOrderService.getCandidates(order.supplier);
}
function addSupplierOrderMaterials(orderId, rowIds) {
  const result = SupplierOrderService.addMaterials(orderId, rowIds);
  return { result: result, snapshot: SupplierOrderService.getSnapshot() };
}
function updateSupplierOrderLine(orderId, lineId, quantity, unit) {
  const result = SupplierOrderService.updateLine(orderId, lineId, quantity, unit);
  return { result: result, snapshot: SupplierOrderService.getSnapshot() };
}
function removeSupplierOrderLines(orderId, lineIds) {
  const result = SupplierOrderService.removeLines(orderId, lineIds);
  return { result: result, snapshot: SupplierOrderService.getSnapshot() };
}
function splitSupplierOrderLines(orderId, lineIds) {
  const result = SupplierOrderService.splitLinesToNewOrder(orderId, lineIds);
  return { result: result, snapshot: SupplierOrderService.getSnapshot() };
}
function setSupplierOrderStatus(orderId, status) {
  const result = SupplierOrderService.setStatus(orderId, status);
  return { result: result, snapshot: SupplierOrderService.getSnapshot() };
}
function saveSupplierOrderInvoiceNumbers(orderId, invoiceNumbers) {
  const result = SupplierOrderService.saveInvoiceNumbers(orderId, invoiceNumbers);
  return { result: result, snapshot: SupplierOrderService.getSnapshot() };
}
function saveSupplierOrderTotal(orderId, amount) {
  const result = SupplierOrderService.saveOrderTotal(orderId, amount);
  return { result: result, snapshot: SupplierOrderService.getSnapshot() };
}
function attachSupplierOrderInvoice(orderId, payload) {
  const result = SupplierInvoiceService.attachQuick(orderId, payload);
  return { result: result, snapshot: SupplierOrderService.getSnapshot() };
}

function createEquipment(data) {
  const result = EquipmentService.create(data);
  return { result: result, snapshot: EquipmentService.getSnapshot() };
}

function updateEquipment(id, data) {
  const result = EquipmentService.update(id, data);
  return { result: result, snapshot: EquipmentService.getSnapshot() };
}

function setEquipmentStatus(id, status) {
  const result = EquipmentService.setStatus(id, status);
  return { result: result, snapshot: EquipmentService.getSnapshot() };
}

function setEquipmentSupplier(id, supplier) {
  const result = EquipmentService.setSupplier(id, supplier);
  return { result: result, snapshot: EquipmentService.getSnapshot() };
}

function getEquipmentProposalFiles(id) {
  return EquipmentService.getProposalFiles(id);
}

function uploadEquipmentProposalFile(id, file) {
  return EquipmentService.uploadProposalFile(id, file);
}

function deleteEquipmentProposalFile(id, fileId) {
  return EquipmentService.deleteProposalFile(id, fileId);
}
function previewEquipmentInvoice(file) { return EquipmentInvoiceService.preview(file); }
function saveEquipmentInvoice(payload) {
  const result = EquipmentInvoiceService.save(payload);
  return { result: result, snapshot: EquipmentInvoiceService.getSnapshot() };
}
function getEquipmentInvoiceDetail(id) { return EquipmentInvoiceService.getInvoiceDetail(id); }
function getEquipmentInvoices(id) { return EquipmentInvoiceService.getEquipmentInvoices(id); }
function markEquipmentInvoicePaid(id, paidAt) {
  const result = EquipmentInvoiceService.markPaid(id, paidAt);
  return { result: result, snapshot: EquipmentInvoiceService.getSnapshot() };
}

function createTabletop(data) { const result = TabletopService.create(data); return { result: result, snapshot: TabletopService.getSnapshot() }; }
function updateTabletop(id, data) { const result = TabletopService.update(id, data); return { result: result, snapshot: TabletopService.getSnapshot() }; }
function setTabletopStatus(id, status) { const result = TabletopService.setStatus(id, status); return { result: result, snapshot: TabletopService.getSnapshot() }; }
function setTabletopSupplier(id, supplier) { const result = TabletopService.setSupplier(id, supplier); return { result: result, snapshot: TabletopService.getSnapshot() }; }
function getTabletopProposalFiles(id) { return TabletopService.getProposalFiles(id); }
function uploadTabletopProposalFile(id, file) { return TabletopService.uploadProposalFile(id, file); }
function deleteTabletopProposalFile(id, fileId) { return TabletopService.deleteProposalFile(id, fileId); }

function getSupplySuppliers() {
  return SupplyService.getSupplierRegistry();
}

function saveSupplySupplier(supplier) {
  return SupplyService.saveSupplierRegistryItem(supplier);
}

function deleteSupplySupplier(name) {
  return SupplyService.deleteSupplierRegistryItem(name);
}

function getConsumableCandidates() {
  return SupplyService.getConsumableCandidates();
}

function saveConsumables(names) {
  const result = SupplyService.saveConsumables(names);
  return { result: { added: result.added }, snapshot: result.snapshot };
}
function saveMaterialRoutes(routes) {
  const result = SupplyService.saveMaterialRoutes(routes);
  return { result: result, snapshot: SupplyService.getSnapshot() };
}
function importExpoCatalog() { return PlateCatalogService.importExpoBatch(); }

function searchSupplyArticles(supplier) {
  const result = SupplyService.searchUnmatchedArticles(supplier);
  return { result: result, snapshot: SupplyService.getSnapshot() };
}

function previewPlateMaterials(files, orderKey) {
  return PlateMaterialService.previewFiles(files, orderKey);
}

function getPlateImportOrders(folderName) {
  return PlateMaterialService.getImportOrders(folderName);
}

function savePlateMaterials(materials, sourceName, orderKey) {
  const result = PlateMaterialService.saveMaterials(materials, sourceName, orderKey);
  return { result: result, snapshot: PlateMaterialService.getSnapshot() };
}

function setPlateOrderSheets(rowId, value) {
  const result = PlateMaterialService.setOrderSheets(rowId, value);
  return { result: result, snapshot: PlateMaterialService.getSnapshot() };
}

function setPlatePurchaseStatus(rowId, status) {
  const result = PlateMaterialService.setStatus(rowId, status);
  return { result: result, snapshot: PlateMaterialService.getSnapshot() };
}

// Служебный повтор только для ранее зафиксированных неудачных уведомлений.
// Интерфейс его не вызывает: запуск выполняется вручную из Apps Script при необходимости.
function retryFailedNotifications(limit) {
  return NotificationService.retryFailedNotifications(limit);
}

// Однократная безопасная подготовка листа адресатов и журнала без отправки писем.
function initializeNotifications() {
  return NotificationService.initialize();
}

// Диагностика справочника адресатов без отправки почты.
function getNotificationRecipientDiagnostics() {
  const result = NotificationService.getRecipientDiagnostics();
  Logger.log("NOTIFICATION_RECIPIENTS " + JSON.stringify(result));
  return result;
}

// Однократный запрос разрешения script.send_mail без фактической отправки письма.
function authorizeNotificationEmail() {
  return NotificationService.requestMailAuthorization();
}

/**
 * Выполняет разрешённое действие над очередью и сразу возвращает её
 * актуальное состояние. Один вызов google.script.run исключает окно, когда
 * действие уже сохранено, а браузер ещё показывает прежний статус.
 */
function getClaimCandidates() {
  return OrderService.getClaimCandidates();
}

function runQueueAction(actionName, args, section, clientVersion) {
  assertClientAppVersion(clientVersion);
  SheetService.resetOrderContexts();
  PerformanceService.start(actionName, section);
  const actions = {
    addOrder: function(order) { return OrderService.createOrder(order); },
    addClaim: function(data) { return OrderService.addClaim(data); },
    moveOrderToDate: function(key, date) { return OrderService.moveOrderToDate(key, date); },
    clearManualOrderDate: function(key) { return OrderService.clearManualOrderDate(key); },
    assignOrderConstructor: function(key, constructor) { return OrderService.assignOrderConstructor(key, constructor); },
    changeOrderDuration: function(key, duration) { return OrderService.changeOrderDuration(key, duration); },
    advanceManagerStage: function(key, stage) { return OrderService.advanceManagerStage(key, stage); },
    reviewByTechnologist: function(key, action, constructor, proformaDays, kdDays) {
      return OrderService.reviewByTechnologist(key, action, constructor, proformaDays, kdDays);
    },
    changeWorkStatus: function(key, action, comment) { return OrderService.changeWorkStatus(key, action, comment); },
    changeConstructorTaskDays: function(key, taskType, days) {
      return OrderService.changeConstructorTaskDays(key, taskType, days);
    },
    changeProductionStatus: function(key, action) { return OrderService.changeProductionStatus(key, action); },
    startPaintingBranch: function(key, branch) { return OrderService.startPaintingBranch(key, branch); },
    sendPaintingBranchToPackaging: function(key, branch) { return OrderService.sendPaintingBranchToPackaging(key, branch); },
    launchManufacturing: function(key) { return OrderService.launchManufacturing(key); },
    changeManufacturingStatus: function(key, branch, action, comment) {
      return OrderService.changeManufacturingStatus(key, branch, action, comment);
    },
    sendSgpToInstallation: function(key, branch) {
      return OrderService.sendSgpToInstallation(key, branch);
    },
    saveInstallationComment: function(key, branch, comment) {
      return OrderService.saveInstallationComment(key, branch, comment);
    },
    saveClaimComment: function(key, comment) { return OrderService.saveClaimComment(key, comment); },
    changeManufacturingStage: function(key, branch, stage) {
      return OrderService.changeManufacturingStage(key, branch, stage);
    },
    moveOrderToOverallStage: function(key, stage) {
      return OrderService.moveOrderToOverallStage(key, stage);
    },
    sendOrderToClaims: function(key, branch, comment) { return OrderService.sendOrderToClaims(key, branch, comment); },
    completeClaim: function(key) { return OrderService.completeClaim(key); },
    restoreCompletedCorpusInstallation: function(key) {
      return OrderService.restoreCompletedCorpusInstallation(key);
    },
    setMdfFacade: function(key, enabled) { return OrderService.setMdfFacade(key, enabled); },
    setVeneerFacade: function(key, enabled) { return OrderService.setVeneerFacade(key, enabled); },
    setFacadeMaterials: function(key, selection) { return OrderService.setFacadeMaterials(key, selection); },
    setCorpusTypes: function(key, selection) { return OrderService.setCorpusTypes(key, selection); },
    setMaterialAreas: function(key, areas) { return OrderService.setMaterialAreas(key, areas); },
    saveProductionDays: function(key, days) { return OrderService.saveProductionDays(key, days); },
    changeSandingStatus: function(key, taskType, action, comment, branch) {
      return OrderService.changeSandingStatus(key, taskType, action, comment, branch);
    },
    changeSandingWorker: function(key, worker) { return OrderService.changeSandingWorker(key, worker); },
    setAvailableSandingWorkers: function(workers) { return OrderService.setAvailableSandingWorkers(workers); },
    changePaintingStatus: function(key, taskType, action, comment, branch) {
      return OrderService.changePaintingStatus(key, taskType, action, comment, branch);
    },
    setProformaApprovedDate: function(key, value) { return OrderService.setProformaApprovedDate(key, value); },
    setOrderDeadlineSettings: function(key, mode, manufacturingDays, manualDate) {
      return OrderService.setOrderDeadlineSettings(key, mode, manufacturingDays, manualDate);
    },
    changeField: function(key, fieldName, value) { return OrderService.updateField(key, fieldName, value); }
  };

  const handler = actions[String(actionName || '')];
  if (!handler) throw new Error('Недопустимое действие очереди.');

  try {
    const actionStartedAt = Date.now();
    const result = handler.apply(null, Array.isArray(args) ? args : []);
    const actionMs = Date.now() - actionStartedAt;
    const patch = result && result.patch;
    if (patch) {
      PerformanceService.finish({ actionMs: actionMs, responseType: "patch", snapshotMs: 0, snapshotOrders: 0, snapshotBytes: 0 });
      return { result: result, patch: patch, appVersion: CONFIG.APP_VERSION, performance: { actionMs: actionMs, responseType: "patch" } };
    }
    const snapshotStartedAt = Date.now();
    const snapshot = OrderService.getOrders(section, SheetService.getActionOrderRows());
    const snapshotMs = Date.now() - snapshotStartedAt;
    const serializationStartedAt = Date.now();
    const snapshotBytes = JSON.stringify(snapshot).length;
    const serializationMs = Date.now() - serializationStartedAt;
    PerformanceService.finish({
      actionMs: actionMs, responseType: "snapshot", snapshotMs: snapshotMs, serializationMs: serializationMs,
      snapshotOrders: (snapshot.orders || []).length, snapshotBytes: snapshotBytes
    });
    snapshot.appVersion = CONFIG.APP_VERSION;
    return { result: result, snapshot: snapshot, appVersion: CONFIG.APP_VERSION, performance: { actionMs: actionMs, snapshotMs: snapshotMs, serializationMs: serializationMs, responseType: "snapshot" } };
  } catch (error) {
    PerformanceService.finish({ failed: true, error: String(error && error.message || error) });
    throw error;
  } finally {
    SheetService.resetOrderContexts();
  }
}

/**
 * Подключение HTML-фрагментов в шаблон Index.html
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Сохраняет новую стадию заказа по номеру договора
 */
function saveStage(contractNumber, newStage) {
  try {
    const ss = SpreadsheetApp.openById(Config.SPREADSHEET_ID);
    const sheet = ss.getSheetByName('Изделия');
    
    if (!sheet) throw new Error('Лист "Изделия" не найден');

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    const contractIndex = headers.indexOf('Договор');
    const stageIndex = headers.indexOf('Этап');

    if (contractIndex === -1 || stageIndex === -1) {
      throw new Error('Не найдены колонки "Договор" или "Этап"');
    }

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][contractIndex]) === String(contractNumber)) {
        sheet.getRange(i + 1, stageIndex + 1).setValue(newStage);
        return { success: true, message: 'Стадия обновлена' };
      }
    }

    throw new Error(`Заказ "${contractNumber}" не найден`);

  } catch (e) {
    Logger.log(`Ошибка saveStage: ${e.message}`);
    throw e;
  }
}

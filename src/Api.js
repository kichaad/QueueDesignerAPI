class Api {

  static handleGet(e) {

    try {

      const resource = e.parameter.resource || "";

      switch (resource) {

        case "orders":
          return Response.success(
            OrderService.getOrders(e.parameter.section || "constructors")
          );

        default:
          return Response.error("Unknown resource");

      }

    } catch (err) {

      return Response.error(err.message);

    }

  }


  static handlePost(e) {

    try {
      const rawBody = e && e.postData && e.postData.contents;
      if (!rawBody || rawBody.length > 14 * 1024 * 1024) {
        throw new Error("BAD_REQUEST|Некорректное тело запроса.");
      }

      let body;
      try {
        body = JSON.parse(rawBody);
      } catch (parseError) {
        throw new Error("BAD_REQUEST|Запрос должен содержать корректный JSON.");
      }

      this.requireTimewebSecret(body.secret);
      const method = String(body.method || "").trim();
      const args = Array.isArray(body.args) ? body.args : [];
      const handlers = this.getTimewebRpcHandlers();
      const handler = handlers[method];
      if (!handler) throw new Error("METHOD_NOT_ALLOWED|Метод API не разрешён.");

      return Response.success(handler.apply(null, args));
    } catch (err) {
      const rawMessage = String(err && err.message || err || "");
      const separator = rawMessage.indexOf("|");
      const code = separator > 0 ? rawMessage.slice(0, separator) : "INTERNAL";
      const message = separator > 0 ? rawMessage.slice(separator + 1) : "Не удалось выполнить запрос.";
      console.error(JSON.stringify({ event: "timeweb_rpc_error", code: code, message: rawMessage }));
      return Response.error(code + "|" + message);
    }

  }

  static requireTimewebSecret(received) {
    const expected = PropertiesService.getScriptProperties().getProperty("TIMEWEB_BACKEND_SECRET");
    if (!expected || !this.constantTimeEquals(expected, String(received || ""))) {
      throw new Error("UNAUTHORIZED|Доступ запрещён.");
    }
  }

  static constantTimeEquals(left, right) {
    const leftDigest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, left, Utilities.Charset.UTF_8);
    const rightDigest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, right, Utilities.Charset.UTF_8);
    let difference = 0;
    for (let index = 0; index < leftDigest.length; index++) {
      difference |= leftDigest[index] ^ rightDigest[index];
    }
    return difference === 0;
  }

  static getTimewebRpcHandlers() {
    return {
      getOrders: getOrders,
      readSupplierOrderTableData: readSupplierOrderTableData,
      readSupplyTableData: readSupplyTableData,
      applyTimewebSheetChanges: packet => TimewebSheetMirror.apply(packet),
      downloadMeasurementArchive: downloadMeasurementArchive,
      getManualRequestChoices: getManualRequestChoices,
      getManualRequestOrderDirectory: getManualRequestOrderDirectory,
      saveManualSupplyRequest: saveManualSupplyRequest,
      exportListExcel: exportListExcel,
      saveSupplierOrderTotal: saveSupplierOrderTotal,
      getClaimCandidates: getClaimCandidates,
      runQueueAction: runQueueAction,
      getPaintingNorms: getPaintingNorms,
      savePaintingNorms: savePaintingNorms,
      activateHourlyPaintingPlanning: activateHourlyPaintingPlanning,
      changeFacadeType: changeFacadeType,
      getOrderFiles: getOrderFiles,
      uploadOrderFile: uploadOrderFile,
      downloadOrderFile: downloadOrderFile,
      getOrderFileThumbnail: getOrderFileThumbnail,
      deleteOrderFile: deleteOrderFile,
      getMailCloudConnectionStatus: getMailCloudConnectionStatus,
      getSupplyImportCandidates: getSupplyImportCandidates,
      importSupplySpecification: importSupplySpecification,
      previewSupplySpecification: previewSupplySpecification,
      analyzeSupplySpecification: analyzeSupplySpecification,
      saveSupplySpecification: saveSupplySpecification,
      changePurchaseStatus: changePurchaseStatus,
      setPurchaseStatus: setPurchaseStatus,
      setSupplyArticle: setSupplyArticle,
      assignSupplySupplier: assignSupplySupplier,
      deleteSupplyMaterial: deleteSupplyMaterial,
      getSupplyOrderDirectory: getSupplyOrderDirectory,
      clearSupplyOrderMaterials: clearSupplyOrderMaterials,
      getSupplierOrderCandidates: getSupplierOrderCandidates,
      createSupplierOrder: createSupplierOrder,
      getSupplierOrderDetails: getSupplierOrderDetails,
      saveSupplierOrderDeliveryDates: saveSupplierOrderDeliveryDates,
      exportSupplierOrderExcel: exportSupplierOrderExcel,
      setSupplierOrderLineArticle: setSupplierOrderLineArticle,
      setSupplierOrderLineMoney: setSupplierOrderLineMoney,
      prepareMdmSupplierOrdersExport: prepareMdmSupplierOrdersExport,
      completeMdmSupplierOrdersExport: completeMdmSupplierOrdersExport,
      exportAllOrdersExcel: exportAllOrdersExcel,
      exportProductionExcel: exportProductionExcel,
      createBasisSpecification: createBasisSpecification,
      getSupplierOrderAddCandidates: getSupplierOrderAddCandidates,
      addSupplierOrderMaterials: addSupplierOrderMaterials,
      updateSupplierOrderLine: updateSupplierOrderLine,
      removeSupplierOrderLines: removeSupplierOrderLines,
      deleteSupplierOrder: deleteSupplierOrder,
      splitSupplierOrderLines: splitSupplierOrderLines,
      setSupplierOrderStatusFast: setSupplierOrderStatusFast,
      saveSupplierOrderInvoiceNumbersFast: saveSupplierOrderInvoiceNumbersFast,
      setSupplierOrderStatus: setSupplierOrderStatus,
      saveSupplierOrderInvoiceNumbers: saveSupplierOrderInvoiceNumbers,
      attachSupplierOrderInvoice: attachSupplierOrderInvoice,
      createEquipment: createEquipment,
      updateEquipment: updateEquipment,
      setEquipmentStatus: setEquipmentStatus,
      setEquipmentSupplier: setEquipmentSupplier,
      getEquipmentProposalFiles: getEquipmentProposalFiles,
      uploadEquipmentProposalFile: uploadEquipmentProposalFile,
      deleteEquipmentProposalFile: deleteEquipmentProposalFile,
      previewEquipmentInvoice: previewEquipmentInvoice,
      saveEquipmentInvoice: saveEquipmentInvoice,
      getEquipmentInvoiceDetail: getEquipmentInvoiceDetail,
      getEquipmentInvoices: getEquipmentInvoices,
      markEquipmentInvoicePaid: markEquipmentInvoicePaid,
      createTabletop: createTabletop,
      updateTabletop: updateTabletop,
      setTabletopStatus: setTabletopStatus,
      setTabletopSupplier: setTabletopSupplier,
      getTabletopProposalFiles: getTabletopProposalFiles,
      uploadTabletopProposalFile: uploadTabletopProposalFile,
      deleteTabletopProposalFile: deleteTabletopProposalFile,
      getSupplySuppliers: getSupplySuppliers,
      saveSupplySupplier: saveSupplySupplier,
      deleteSupplySupplier: deleteSupplySupplier,
      getConsumableCandidates: getConsumableCandidates,
      saveConsumables: saveConsumables,
      saveMaterialRoutes: saveMaterialRoutes,
      importExpoCatalog: importExpoCatalog,
      searchSupplyArticles: searchSupplyArticles,
      previewPlateMaterials: previewPlateMaterials,
      getPlateImportOrders: getPlateImportOrders,
      savePlateMaterials: savePlateMaterials,
      setPlateOrderSheets: setPlateOrderSheets,
      setPlatePurchaseStatus: setPlatePurchaseStatus
    };
  }

}

function setupTimewebBackendSecret() {
  const secret = Utilities.getUuid().replace(/-/g, "") + Utilities.getUuid().replace(/-/g, "");
  PropertiesService.getScriptProperties().setProperty("TIMEWEB_BACKEND_SECRET", secret);
  return secret;
}

/**
 * Универсальное изменение поля заказа
 * Используется клиентским приложением через google.script.run
 */
function changeField(
    key,
    fieldName,
    value
){

    return OrderService.updateField(
        key,
        fieldName,
        value
    );

}

/**
 * Staged SQL→Google receiver for the main "Изделия" sheet.
 *
 * DISABLED until production order creation, constructor scheduling and
 * management migrate to SQL. Do not enable while Google remains a writer.
 * Called only through authenticated Timeweb RPC; never writes back to SQL.
 */
class TimewebProductionMirror {
  static apply(packet) {
    if (PropertiesService.getScriptProperties().getProperty("TIMEWEB_PRODUCTION_MIRROR_ENABLED") !== "enabled") {
      throw new Error("SYNC_DISABLED|Зеркалирование производственных заказов не включено.");
    }
    if (!packet || packet.schema !== 1 || !Number.isSafeInteger(packet.productionOrderId)
        || packet.productionOrderId <= 0 || packet.event !== "upsert"
        || !packet.fields || typeof packet.fields !== "object") {
      throw new Error("BAD_REQUEST|Некорректная запись производственного заказа.");
    }
    const allowed = ["Договор","Клиент","Изделие","Менеджер"];
    if (Object.keys(packet.fields).some(name => !allowed.concat("ID заказа на производство").includes(name))
        || !allowed.every(name => typeof packet.fields[name] === "string")
        || allowed.some(name => packet.fields[name].length > 255)
        || !packet.fields["Клиент"].trim() || !packet.fields["Изделие"].trim()
        || !packet.fields["Менеджер"].trim()
        || Number(packet.fields["ID заказа на производство"]) !== packet.productionOrderId) {
      throw new Error("BAD_REQUEST|Не хватает реквизитов заказа для очереди конструкторов.");
    }
    const sheet = SheetService.getSpreadsheet().getSheetByName(CONFIG.SHEETS.ORDERS);
    if (!sheet || sheet.getLastRow() < 1) {
      throw new Error("SYNC_SCHEMA|Основной лист «Изделия» отсутствует.");
    }
    // Manual verification of duplicate IDs and legacy identities takes
    // place before any column is created or any data is changed.
    const values = sheet.getDataRange().getValues();
    const headers = values[0].map(x => String(x || "").trim());
    const index = name => headers.indexOf(name);
    if (["Договор","Клиент","Изделие"].some(name => index(name) < 0)) {
      throw new Error("SYNC_SCHEMA|Лист «Изделия» не содержит составной ключ.");
    }
    const orderIdCol = index("ID заказа на производство");
    const key = [packet.fields["Договор"],packet.fields["Клиент"],packet.fields["Изделие"]].join("|");
    const byId = [], byKey = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (orderIdCol >= 0 && String(row[orderIdCol] || "") === String(packet.productionOrderId)) byId.push(i+1);
      if (["Договор","Клиент","Изделие"].map(name => String(row[index(name)] || "")).join("|") === key) byKey.push(i+1);
    }
    if (byId.length > 1 || byKey.length > 1 || (byId.length && byKey.length && byId[0] !== byKey[0])) {
      throw new Error("SYNC_CONFLICT|Неоднозначный ID или дублирующий бизнес-ключ.");
    }
    if (byId.length && !byKey.length) {
      throw new Error("SYNC_CONFLICT|В Google изменён составной ключ заказа.");
    }
    const existingRow = byId[0] || byKey[0];
    if (!existingRow) {
      // The existing appender sets defaults, queue position and formulas.
      // It uses its own lock, and checks uniqueness under that lock.
      SheetService.appendOrder({
        contract:packet.fields["Договор"], client:packet.fields["Клиент"],
        product:packet.fields["Изделие"], manager:packet.fields["Менеджер"],
        constructor:"", duration:1,
      });
    }
    const target = sheet;
    const idSheet = SheetService.ensureOrderColumn("ID заказа на производство");
    const currentHeaders=idSheet.getRange(1,1,1,idSheet.getLastColumn()).getValues()[0].map(x=>String(x||"").trim());
    const columns=currentHeaders.map((x,i)=>[x,i+1]);
    const idColumn=columns.find(entry=>entry[0]==="ID заказа на производство")[1];
    const legacyColumns=["Договор","Клиент","Изделие"].map(n=>currentHeaders.indexOf(n)+1);
    const rows=idSheet.getDataRange().getValues();
    const matches=[];
    rows.forEach((row,i)=> {
      if(i>0 && legacyColumns.every((col,j)=>String(row[col-1] || "")===packet.fields[["Договор","Клиент","Изделие"][j]]))matches.push(i+1);
    });
    if(matches.length!==1)throw new Error("SYNC_CONFLICT|Не удалось однозначно найти строку заказа после создания.");
    const currentId=rows[matches[0]-1][idColumn-1];
    if(currentId!=="" && currentId!==null && String(currentId)!==String(packet.productionOrderId)) {
      throw new Error("SYNC_CONFLICT|Строка уже привязана к другому производственному заказу.");
    }
    idSheet.getRange(matches[0],idColumn).setValue(packet.productionOrderId);
    return {productionOrderId:packet.productionOrderId,existing:Boolean(existingRow)};
  }
}

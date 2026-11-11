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
    const required=["Договор","Клиент","Изделие","Менеджер"];
    const allowed=new Set(required.concat([
      "ID заказа на производство","Этап","Статус","Ответственный",
      "Конструктор","Очередность","Очередь","Длительность",
      "Начало","Окончание","Плановая дата","Дни Проформы","Дни КД",
      "Резерв КД — очередь","Ручной приоритет даты",
      "Факт. начало работы","Факт. половина начала","Остаток длительности"
    ]));
    if (Object.keys(packet.fields).some(name => !allowed.has(name))
        || !required.every(name => typeof packet.fields[name] === "string")
        || required.some(name => packet.fields[name].length > 255)
        || Object.entries(packet.fields).some(([name,value])=>
             name!=="ID заказа на производство" &&
             (value===null || !["string","number","boolean"].includes(typeof value)
              || (typeof value==="string"&&value.length>1000)
              || (typeof value==="number"&&!Number.isFinite(value))))
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
    if (byId.length > 1 || (!byId.length && byKey.length > 1) || (byId.length && byKey.length && !byKey.includes(byId[0]))) {
      throw new Error("SYNC_CONFLICT|Неоднозначный ID или дублирующий бизнес-ключ.");
    }
    // A stable numeric ID permits a later SQL-authoritative rename.
    // If the new legacy key belongs to some OTHER row, reject the conflict.
    const existingRow = byId[0] || byKey[0];
    if (!existingRow) {
      // The existing appender sets defaults, queue position and formulas.
      // It uses its own lock, and checks uniqueness under that lock.
      // Only the authenticated SQL mirror may append after Google writes
      // are fenced. Restore the guard in finally, including on exceptions.
      this._writing=true;
      try {
        SheetService.appendOrder({
          contract:packet.fields["Договор"], client:packet.fields["Клиент"],
          product:packet.fields["Изделие"], manager:packet.fields["Менеджер"],
          constructor:"", duration:0.5,
        });
      } finally { this._writing=false; }
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
    // Existing immutable SQL ID disambiguates historical duplicate names.
    const chosen=byId.length===1?byId[0]:(matches.length===1?matches[0]:null);
    if(!chosen)
      throw new Error("SYNC_CONFLICT|Не удалось однозначно найти строку заказа после создания.");
    const currentId=rows[chosen-1][idColumn-1];
    if(currentId!=="" && currentId!==null && String(currentId)!==String(packet.productionOrderId)) {
      throw new Error("SYNC_CONFLICT|Строка уже привязана к другому производственному заказу.");
    }
    // Only SQL-owned fields are patched; unrelated formulas, cell styles
    // and reports are untouched. Retries write absolute identical values.
    const updated=[]; const missing=[];
    Object.entries(packet.fields).forEach(([name,value])=>{
      const col=currentHeaders.indexOf(name);
      if(col<0){missing.push(name);return;}
      const literal=typeof value==="string" && /^[=']/u.test(value) ? "'"+value : value;
      if(String(rows[chosen-1][col]??'')!==String(literal)){
        idSheet.getRange(chosen,col+1).setValue(literal);
        updated.push(name);
      }
    });
    return {productionOrderId:packet.productionOrderId,existing:Boolean(existingRow),
      updatedFields:updated.length,skippedMissingColumns:missing.filter(x=>x!=="ID заказа на производство")};
  }
}

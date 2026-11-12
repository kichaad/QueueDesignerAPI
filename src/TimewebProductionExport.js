/**
 * Explicit read-only, paginated one-time export for SQL production migration.
 * OFF by default, authenticated by the existing Timeweb RPC secret.
 */
class TimewebProductionExport {
  static page(offset,limit) {
    if (PropertiesService.getScriptProperties().getProperty("TIMEWEB_PRODUCTION_EXPORT_ENABLED") !== "enabled") {
      throw new Error("EXPORT_DISABLED|Экспорт для миграции не активирован.");
    }
    const start=Number(offset), size=Number(limit);
    if(!Number.isInteger(start)||start<0||!Number.isInteger(size)||size<1||size>400)
      throw new Error("BAD_REQUEST|Некорректный диапазон экспорта.");
    const sheet=SheetService.getSpreadsheet().getSheetByName(CONFIG.SHEETS.ORDERS);
    if(!sheet)throw new Error("SYNC_SCHEMA|Лист «Изделия» не найден.");
    const width=sheet.getLastColumn();
    const count=Math.max(0,sheet.getLastRow()-1);
    if(width<1||width>250||count>20000)throw new Error("BAD_REQUEST|Лист слишком велик.");
    const headers=sheet.getRange(1,1,1,width).getDisplayValues()[0].map(x=>String(x||"").trim());
    const amount=Math.min(size,Math.max(0,count-start));
    // getDisplayValues avoids serializing GAS Date objects or formulas. The
    // original sheet stays unchanged. Dates are exported as displayed strings.
    const rows=amount
      ? sheet.getRange(start+2,1,amount,width).getDisplayValues()
      : [];
    return {schema:1,sheet:"Изделия",headers,rows,offset:start,total:count,
      nextOffset:start+amount<count?start+amount:null};
  }
}

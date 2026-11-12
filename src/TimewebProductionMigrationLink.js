/**
 * One-time attachment of SQL production_order_id to frozen Google source rows.
 * Requires explicit migration flag AND disabled Google operational writes.
 * Never updates the contract/client/product/stage columns themselves.
 */
class TimewebProductionMigrationLink {
  static apply(mappings) {
    const props=PropertiesService.getScriptProperties();
    if(props.getProperty('TIMEWEB_PRODUCTION_LINK_ENABLED')!=='enabled'
        ||(props.getProperty('TIMEWEB_PRODUCTION_GOOGLE_WRITES_DISABLED')!=='enabled'&&!(typeof TimewebConstructorSql!=='undefined'&&TimewebConstructorSql.frozen())))
      throw new Error('MIGRATION_NOT_FROZEN|Google queue must be frozen before mapping SQL IDs.');
    if(!Array.isArray(mappings)||!mappings.length||mappings.length>100)
      throw new Error('BAD_REQUEST|ID mapping must have between 1 and 100 items.');
    const lock=SheetService.productionLock();
    lock.waitLock(30000);
    try {
      const sheet=SheetService.getSpreadsheet().getSheetByName(CONFIG.SHEETS.ORDERS);
      if(!sheet)throw new Error('SYNC_SCHEMA|Missing production source sheet.');
      let values=sheet.getDataRange().getValues();
      const headers=values[0].map(x=>String(x||'').trim());
      const ix=name=>headers.indexOf(name);
      if(['Договор','Клиент','Изделие','Менеджер','Этап'].some(name=>ix(name)<0))
        throw new Error('SYNC_SCHEMA|Required source columns missing.');
      const idCol=ix('ID заказа на производство');
      if(idCol>=0&&sheet.getLastRow()>1){sheet.getRange(2,idCol+1,sheet.getLastRow()-1,1).setNumberFormat('0');values=sheet.getDataRange().getValues();}
      const used=new Map();const targets=[];
      if(idCol>=0){
        for(let row=1;row<values.length;row++){
          const current=values[row][idCol];
          if(current!==''&&current!==null&&current!==undefined){
            const key=String(current);
            if(used.has(key)&&used.get(key)!==row+1)throw new Error('SYNC_CONFLICT|Duplicate SQL ID in Google.');
            used.set(key,row+1);
          }
        }
      }
      mappings.forEach(item=>{
        const row=Number(item.sheetRow),id=Number(item.productionOrderId);
        if(!Number.isInteger(row)||row<2||row>values.length||
           !Number.isSafeInteger(id)||id<=0 ||typeof item.key!=='string'||
           typeof item.stage!=='string'||typeof item.manager!=='string')
          throw new Error('BAD_REQUEST|Invalid SQL ID mapping.');
        const source=values[row-1];
        const key=['Договор','Клиент','Изделие']
          .map(name=>String(source[ix(name)]||'').trim()).join('|');
        if(key!==item.key || String(source[ix('Этап')]||'').trim()!==item.stage ||
           String(source[ix('Менеджер')]||'').trim()!==item.manager)
          throw new Error('SYNC_CONFLICT|Source production row changed after export.');
        if(used.has(String(id))&&used.get(String(id))!==row)
          throw new Error('SYNC_CONFLICT|SQL production ID is already bound to a different row.');
        if(idCol>=0){
          const previous=source[idCol];
          if(previous!=='' && previous!==null && previous!==undefined && String(previous)!==String(id))
            throw new Error('SYNC_CONFLICT|Google row already has a different SQL ID.');
        }
        if(targets.some(x=>x.row===row && x.id!==id))
          throw new Error('BAD_REQUEST|Conflicting mappings for one row.');
        used.set(String(id),row);
        targets.push({row,id});
      });
      const targetSheet=idCol<0?SheetService.ensureOrderColumn('ID заказа на производство'):sheet;
      const column=idCol<0?targetSheet.getLastColumn():idCol+1;
      targets.forEach(item=>targetSheet.getRange(item.row,column).setNumberFormat('0').setValue(item.id));
      return {tagged:targets.length};
    } finally {lock.releaseLock();}
  }
}

/** Authenticated control plane and idempotent report/email delivery for the SQL constructor master. */
class TimewebConstructorSql {
  static stages() { return ['ТЗ','Технолог','Проформа','Согласование','Правки','КД']; }
  static frozen() { return PropertiesService.getScriptProperties().getProperty('TIMEWEB_CONSTRUCTOR_SQL_ENABLED')==='enabled'; }
  static dictionary() {
    const sheet=SheetService.getSpreadsheet().getSheetByName(CONFIG.SHEETS.DICTIONARIES);
    if(!sheet)return {};
    const rows=sheet.getDataRange().getDisplayValues(),out={};
    (rows[0]||[]).forEach((h,c)=>{h=String(h||'').trim();if(h)out[h]=rows.slice(1).map(r=>r[c]).filter(x=>x!==''&&x!=null);});
    return out;
  }
  static manage(action,payload) {
    const lock=SheetService.productionLock();lock.waitLock(30000);
    try {
      const props=PropertiesService.getScriptProperties();
      if(action==='freeze'){
        props.setProperty('TIMEWEB_CONSTRUCTOR_SQL_ENABLED','enabled');
        props.setProperty('TIMEWEB_PRODUCTION_LINK_ENABLED','enabled');
      }else if(action==='thaw'){
        if(!payload||payload.confirm!=='no-sql-writes')throw new Error('BAD_REQUEST|SQL write reconciliation required before rollback.');
        props.deleteProperty('TIMEWEB_CONSTRUCTOR_SQL_ENABLED');
        return {frozen:false};
      }else if(!['status','export','link','reference'].includes(action))throw new Error('BAD_REQUEST|Unknown SQL constructor control action.');
      if(action==='link')return TimewebProductionMigrationLink.apply(payload);
      const sheet=SheetService.getSpreadsheet().getSheetByName(CONFIG.SHEETS.ORDERS);
      const rows=sheet.getDataRange().getDisplayValues(),headers=(rows[0]||[]).map(x=>String(x||'').trim());
      const stage=headers.indexOf('Этап'),client=headers.indexOf('Клиент'),product=headers.indexOf('Изделие');
      const count=rows.slice(1).filter(r=>String(r[client]||'').trim()||String(r[product]||'').trim()).length;
      if(action==='status'||action==='freeze')return {frozen:this.frozen(),rows:count,
        designRows:rows.slice(1).filter(r=>this.stages().includes(String(r[stage]||'').trim())).length};
      if(!this.frozen())throw new Error('MIGRATION_NOT_FROZEN|Freeze constructor writers before export.');
      const offset=Number(payload&&payload.offset||0),limit=Number(payload&&payload.limit||200);
      if(!Number.isInteger(offset)||offset<0||!Number.isInteger(limit)||limit<1||limit>400||rows.length>20001||headers.length>250)
        throw new Error('BAD_REQUEST|Invalid export range.');
      const raw=sheet.getDataRange().getValues(),effectiveStages={};
      const mapped=SheetService.objectsFromRows(raw).map((r,i)=>{
        const order=OrderMapper.map(r),overall=OrderService.resolveOverallStage(order);
        if(overall){order.stage=overall.stage;order.status=overall.status;}
        effectiveStages[i+2]=order.stage;
        const id=rows[i+1][headers.indexOf('ID заказа на производство')];
        if(id)order.key='sql-production:'+id;
        return order;
      });
      if(action==='reference'){
        const tasks=OrderService.getConstructorTasks(mapped);CalendarService.buildCalendar(tasks);
        return {today:CalendarService.formatDate(CalendarService.firstWorkDay()),half:CalendarService.getCurrentHalfDay(),
          tasks:tasks.map(t=>({key:t.key,taskType:t.taskType,startDate:t.startDate,finishDate:t.finishDate,queue:t.queue,duration:t.duration,dayPart:t.dayPart||''})),
          designOrders:mapped.filter(o=>this.stages().includes(o.stage)).length};
      }
      const data=rows.slice(1),page=data.slice(offset,offset+limit);
      return {schema:1,sheet:'Изделия',headers,rows:page,total:data.length,offset,
        nextOffset:offset+page.length<data.length?offset+page.length:null,dictionaries:this.dictionary(),effectiveStages};
    }finally{lock.releaseLock();}
  }
  static apply(packet) {
    if(!this.frozen())throw new Error('SQL_DISABLED|SQL constructor master is not enabled.');
    if(!packet||packet.schema!==1||!Number.isSafeInteger(packet.eventId)||packet.eventId<1
        ||!Number.isSafeInteger(packet.productionOrderId)||packet.productionOrderId<1||!packet.fields)
      throw new Error('BAD_REQUEST|Invalid constructor event.');
    const allowed=new Set(['ID заказа на производство','Договор','Клиент','Изделие','Менеджер','Этап','Статус','Ответственный','Конструктор',
      'Очередность','Длительность','Начало','Окончание','Плановая дата','Дни ТЗ','Дни Проформы','Дни КД','Резерв КД — очередь',
      'Ручной приоритет даты','Факт. начало работы','Факт. половина начала','Остаток длительности','Комментарий',
      'Проформа Согласована','Способ определения дедлайна','Срок изготовления, дней','Дата дедлайна вручную','Очередь производства','Папка Google Drive']);
    if(Object.entries(packet.fields).some(([k,v])=>!allowed.has(k)||!['number','string','boolean'].includes(typeof v)
        ||(typeof v==='string'&&v.length>2000)||(typeof v==='number'&&!Number.isFinite(v)))
        ||Number(packet.fields['ID заказа на производство'])!==packet.productionOrderId)
      throw new Error('BAD_REQUEST|Invalid constructor report fields.');
    const lock=SheetService.productionLock();lock.waitLock(30000);
    try {
      const props=PropertiesService.getScriptProperties(),receipt='SQL_CTOR_DELIVERED_'+packet.productionOrderId;
      if(Number(props.getProperty(receipt)||0)>=packet.eventId)return {eventId:packet.eventId,productionOrderId:packet.productionOrderId,duplicate:true};
      const sheet=SheetService.ensureOrderColumn('ID заказа на производство');
      let values=sheet.getDataRange().getValues(),headers=values[0].map(x=>String(x||'').trim());
      // Add only explicitly owned columns. Never replace unrelated report/formula cells.
      Object.keys(packet.fields).forEach(k=>{if(!headers.includes(k))SheetService.ensureOrderColumn(k);});
      values=sheet.getDataRange().getValues();headers=values[0].map(x=>String(x||'').trim());
      const idCol=headers.indexOf('ID заказа на производство'),keyCols=['Договор','Клиент','Изделие'].map(k=>headers.indexOf(k));
      const matches=[];for(let i=1;i<values.length;i++)if(String(values[i][idCol]||'')===String(packet.productionOrderId))matches.push(i+1);
      if(matches.length>1)throw new Error('SYNC_CONFLICT|Duplicate production ID.');
      let row=matches[0];
      if(!row){
        if(values.slice(1).some(r=>keyCols.every((c,j)=>String(r[c]||'')===String(packet.fields[['Договор','Клиент','Изделие'][j]]||''))))
          throw new Error('SYNC_CONFLICT|Existing legacy order has no matching SQL ID.');
        const free=values.findIndex((r,i)=>i>0&&keyCols.every(c=>String(r[c]||'').trim()===''));
        row=free>=0?free+1:sheet.getLastRow()+1;
        if(row>sheet.getMaxRows())sheet.insertRowAfter(sheet.getMaxRows());
      }
      Object.entries(packet.fields).forEach(([k,v])=>{
        const literal=typeof v==='string'&&/^[=']/u.test(v)?"'"+v:v;
        sheet.getRange(row,headers.indexOf(k)+1).setValue(literal);
      });
      SpreadsheetApp.flush();
      if(packet.notification){
        const n=packet.notification;
        if(!NotificationService.getEventDefinition(n.type))throw new Error('BAD_REQUEST|Unknown notification event.');
        const result=NotificationService.onAssignmentChanged(n.before,n.after,{eventType:n.type});
        const statuses=Array.isArray(result)?result:[result];
        if(statuses.some(r=>r&&r.status==='FAILED'))throw new Error('NOTIFICATION_FAILED|Notification will be retried.');
      }
      props.setProperty(receipt,String(packet.eventId));
      return {eventId:packet.eventId,productionOrderId:packet.productionOrderId};
    }finally{lock.releaseLock();}
  }
  static fileAction(method,order,args) {
    const names=['getOrderFiles','uploadOrderFile','downloadOrderFile','downloadMeasurementArchive','getOrderFileThumbnail','deleteOrderFile'];
    if(!names.includes(method)||!order||!/^sql-production:[1-9][0-9]*$/.test(order.key)||!Array.isArray(args))throw new Error('BAD_REQUEST|Invalid SQL file operation.');
    // File storage receives SQL-owned identity and metadata; no Google order lookup.
    const fn={getOrderFiles:'getOrderFiles',uploadOrderFile:'uploadOrderFile',downloadOrderFile:'downloadOrderFile',
      downloadMeasurementArchive:'downloadMeasurementArchive',getOrderFileThumbnail:'getOrderFileThumbnail',deleteOrderFile:'deleteOrderFile'}[method];
    order.sqlOwned=true;
    const result=GoogleDriveService[fn].apply(GoogleDriveService,[order].concat(args));
    return {result,cloudFolderUrl:order.cloudFolderUrl||''};
  }
}

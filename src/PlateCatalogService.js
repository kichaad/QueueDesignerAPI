/** Справочник плит и декоров Expo-Торг. */
class PlateCatalogService {
  static get HEADERS() { return ["Категория", "Бренд", "Код декора", "Декор", "Толщина, мм", "Ссылка", "Обновлено"]; }
  static get SOURCES() { return [["ЛДСП", "https://expo-torg.ru/catalog/plity-s-pokrytiem/ldsp/egger/"], ["ЛМДФ", "https://expo-torg.ru/catalog/plity-s-pokrytiem/lmdf/"], ["ЛХДФ", "https://expo-torg.ru/catalog/plity-s-pokrytiem/lkhdf/"], ["МДФ", "https://expo-torg.ru/catalog/plity-bez-pokrytiya/mdf/"]]; }

  static ensureSheet() {
    const book = SheetService.getSpreadsheet(); let sheet = book.getSheetByName(CONFIG.SHEETS.PLATE_CATALOG);
    if (!sheet) sheet = book.insertSheet(CONFIG.SHEETS.PLATE_CATALOG);
    if (!sheet.getLastRow()) { sheet.getRange(1, 1, 1, this.HEADERS.length).setValues([this.HEADERS]); sheet.setFrozenRows(1); sheet.getRange(1, 1, 1, this.HEADERS.length).setBackground("#155e75").setFontColor("#fff").setFontWeight("bold"); sheet.getRange(1, 1, 1, this.HEADERS.length).createFilter(); }
    return sheet;
  }
  static getIndex() { return this.ensureSheet().getDataRange().getDisplayValues().slice(1).filter(row => row[0]).map(row => ({ category: row[0], brand: row[1], code: row[2], decor: row[3], thickness: row[4], url: row[5] })); }
  static match(material) {
    const value = String(material.name || "") + " " + String(material.article || "");
    const category = /лдсп/i.test(value) ? "ЛДСП" : /лмдф/i.test(value) ? "ЛМДФ" : /лхдф/i.test(value) ? "ЛХДФ" : /мдф/i.test(value) ? "МДФ" : /хдф/i.test(value) ? "ЛХДФ" : "";
    const code = (value.match(/\b(?:[A-Z]{1,3}\d{2,4}(?:\s*[A-Z]{1,3}\d{0,3})?|\d{3,5})\b/i) || [])[1] || "";
    const thickness = (value.match(/\b(\d{1,2})\s*мм/i) || [])[1] || "";
    const exact = this.getIndex().find(item => item.category === category && code && SupplyService.normalize(item.code) === SupplyService.normalize(code));
    return { category: category, code: exact ? exact.code : code, decor: exact ? `${exact.code} ${exact.decor}`.trim() : "", thickness: exact ? exact.thickness : thickness };
  }

  // Берём текст только из ссылок-карточек товаров. Прежний поиск по всей
  // «склеенной» HTML-странице давал случайные 1–2 совпадения.
  static extractProductRows(html, category, baseUrl) {
    const rows = [], anchor = /<a\b([^>]*\bhref\s*=\s*["']([^"']+)["'][^>]*)>([\s\S]*?)<\/a>/gi;
    let match;
    while ((match = anchor.exec(String(html || "")))) {
      const title = this.cleanHtml(match[3]);
      if (!new RegExp(`^${category}\\b`, "i").test(title) || !/\b\d{2,4}\s*[*xх×]\s*\d{3,4}\s*[*xх×]\s*\d{3,4}\b/i.test(title)) continue;
      const parsed = this.parseProductTitle(title, category);
      if (parsed) rows.push([category, parsed.brand, parsed.code, parsed.decor, parsed.thickness, this.absoluteUrl(match[2], baseUrl), new Date()]);
    }
    return rows;
  }
  static cleanHtml(value) { return String(value || "").replace(/<[^>]+>/g, " ").replace(/&nbsp;|&#160;/gi, " ").replace(/&quot;/gi, "\"").replace(/&amp;/gi, "&").replace(/&#(\d+);/g, (_, code) => String.fromCharCode(Number(code))).replace(/&#x([0-9a-f]+);/gi, (_, code) => String.fromCharCode(parseInt(code, 16))).replace(/\s+/g, " ").trim(); }
  static absoluteUrl(url, baseUrl) { if (/^https?:\/\//i.test(url)) return url; return url.charAt(0) === "/" ? "https://expo-torg.ru" + url : baseUrl.replace(/\/[^/]*$/, "/") + url; }
  static parseProductTitle(title, category) {
    let rest = String(title).replace(new RegExp(`^${category}\\s*`, "i"), "").trim();
    const codeMatch = rest.match(/\b(?:[A-Z]{1,3}\s*\d{2,4}|\d{3,5})\b/i), code = codeMatch ? codeMatch[0].replace(/\s+/g, "") : "";
    const sizeMatch = rest.match(/\b(0?\d{1,2})\s*[*xх×]\s*\d{3,4}\s*[*xх×]\s*\d{3,4}\s*мм?\b/i);
    if (!sizeMatch) return null;
    const thickness = String(Number(sizeMatch[1])), brandMatch = rest.match(/\b(Egger|Кроношпан Беларусь|Кроношпан|Kronospan)\b/i), brand = brandMatch ? brandMatch[0] : (category === "ЛДСП" ? "Egger" : "");
    let decor = rest.replace(/\b(Egger|Кроношпан Беларусь|Кроношпан|Kronospan)\b/ig, " ").replace(/\b(?:PE|ST\s*\d{1,3})\b/ig, " ").replace(/\b0?\d{1,2}\s*[*xх×]\s*\d{3,4}\s*[*xх×]\s*\d{3,4}\s*мм?\b/ig, " ").replace(/\b(?:одно|двух|1|2)[-\s]*сторонн\w*\b/ig, " ");
    if (code) decor = decor.replace(new RegExp(`\\b${code.replace(/[.*+?^${}()|[\]\\]/g, "\\$&")}\\b`, "i"), " ");
    decor = decor.replace(/\s+/g, " ").replace(/^[,;\-–—]+|[,;\-–—]+$/g, "").trim();
    return { code: code, decor: decor || title, thickness: thickness, brand: brand };
  }
  static clearExpoRows(sheet) {
    const values = sheet.getDataRange().getValues(), keep = values.slice(1).filter(row => !/expo-torg\.ru/i.test(String(row[5] || "")));
    sheet.clearContents(); sheet.getRange(1, 1, 1, this.HEADERS.length).setValues([this.HEADERS]);
    if (keep.length) sheet.getRange(2, 1, keep.length, this.HEADERS.length).setValues(keep);
  }
  static importExpoBatch() {
    const props = PropertiesService.getScriptProperties(); let state = JSON.parse(props.getProperty("plateCatalogImport") || "null");
    if (state && state.source >= this.SOURCES.length) { props.deleteProperty("plateCatalogImport"); state = null; }
    const sheet = this.ensureSheet();
    if (!state) { this.clearExpoRows(sheet); state = { source: 0, page: 1 }; }
    const [category, baseUrl] = this.SOURCES[state.source], url = state.page === 1 ? baseUrl : `${baseUrl}?PAGEN_1=${state.page}`;
    const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true, followRedirects: true }), html = response.getContentText();
    if (response.getResponseCode() >= 400) throw new Error(`Expo-Торг вернул HTTP ${response.getResponseCode()} для ${category}.`);
    const rows = this.extractProductRows(html, category, baseUrl), existing = new Set(sheet.getDataRange().getDisplayValues().slice(1).map(row => [row[0], row[2], row[3], row[4]].map(SupplyService.normalize).join("|")));
    const unique = rows.filter(row => { const key = [row[0], row[2], row[3], row[4]].map(SupplyService.normalize).join("|"); if (existing.has(key)) return false; existing.add(key); return true; });
    if (unique.length) sheet.getRange(sheet.getLastRow() + 1, 1, unique.length, this.HEADERS.length).setValues(unique);
    const hasNext = new RegExp(`[?&]PAGEN_1=${state.page + 1}(?:[&"'])`, "i").test(html);
    if (hasNext && rows.length) state.page += 1; else { state.source += 1; state.page = 1; }
    const done = state.source >= this.SOURCES.length;
    if (done) props.deleteProperty("plateCatalogImport"); else props.setProperty("plateCatalogImport", JSON.stringify(state));
    return { done: done, added: unique.length, category: category, page: state.page, found: rows.length };
  }
}

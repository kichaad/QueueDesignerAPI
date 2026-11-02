class LoggerService {

  static info(message) {
    Logger.log(message);
  }

  static error(error) {
    Logger.log(error);
  }

}

/** Короткоживущий профиль одного RPC-вызова Apps Script. */
class PerformanceService {
  static start(action, section) {
    this.current = {
      action: String(action || ""), section: String(section || ""), startedAt: Date.now(),
      lookupMs: 0, sheetReadMs: 0, sheetWriteMs: 0, sheetReads: 0, sheetWrites: 0
    };
    return this.current;
  }

  static measureRead(callback) {
    const startedAt = Date.now();
    try { return callback(); }
    finally {
      if (this.current) {
        this.current.sheetReadMs += Date.now() - startedAt;
        this.current.sheetReads += 1;
      }
    }
  }

  static measureWrite(callback) {
    const startedAt = Date.now();
    try { return callback(); }
    finally {
      if (this.current) {
        this.current.sheetWriteMs += Date.now() - startedAt;
        this.current.sheetWrites += 1;
      }
    }
  }

  static finish(extra) {
    if (!this.current) return;
    const profile = Object.assign({}, this.current, extra || {}, { totalMs: Date.now() - this.current.startedAt });
    Logger.log("QUEUE_PERF " + JSON.stringify(profile));
    this.current = null;
  }
}

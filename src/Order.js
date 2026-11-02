/**
 * Доменная модель "Изделие"
 */
class Order {

  constructor(data) {

    this.key = data.key;

    this.manager = data.manager;

    this.contract = data.contract;

    this.client = data.client;

    this.product = data.product;

    this.displayName = data.displayName;

    this.stage = data.stage;

    this.responsible = data.responsible;

    this.constructor = data.constructor;

    this.queue = data.queue;

    this.startDate = data.startDate;

    this.plannedDate = data.plannedDate;

    this.deadline = data.deadline;

    this.deadlineMode = data.deadlineMode || "От согласования проформы";

    this.manufacturingDays = data.manufacturingDays || 50;

    this.manualDeadlineDate = data.manualDeadlineDate || "";

    this.manualPriority = Boolean(data.manualPriority);

    this.actualStartDate = data.actualStartDate;

    this.actualStartHalf = data.actualStartHalf;

    this.remainingDuration = data.remainingDuration || 0;

    this.duration = data.duration;

    this.tzDays = data.tzDays;

    this.proformaDays = data.proformaDays;

    this.kdDays = data.kdDays;

    this.kdReserveQueue = data.kdReserveQueue;

    this.proformaApprovedDate = data.proformaApprovedDate || "";

    this.finishDate = data.finishDate;

    this.comment = data.comment;

    this.installationCorpusComment = data.installationCorpusComment || "";
    this.installationFacadeComment = data.installationFacadeComment || "";
    this.installationMdfComment = data.installationMdfComment || "";
    this.installationVeneerComment = data.installationVeneerComment || "";
    this.installationLegacyComment = data.installationLegacyComment || "";
    this.claimComment = data.claimComment || "";
    this.manualClaim = Boolean(data.manualClaim);

    this.status = data.status;

    this.cloudFolderUrl = data.cloudFolderUrl;

    this.productionDays = data.productionDays;

    this.lacquerOperations = Boolean(data.lacquerOperations);

    this.routeExclusions = data.routeExclusions || {};

    this.paintingArea = data.paintingArea || 0;

    this.backPaintingArea = data.backPaintingArea || 0;

    this.frontPaintingArea = data.frontPaintingArea || 0;

    this.polishingArea = data.polishingArea || 0;

    this.ldspArea = data.ldspArea || 0;
    this.mdfArea = data.mdfArea || 0;
    this.lhdfArea = data.lhdfArea || 0;
    this.veneerArea = data.veneerArea || 0;

    this.facadeType = data.facadeType || "";

    this.paintingOperationCode = data.paintingOperationCode || "";

    this.productionQueue = data.productionQueue;

    this.paintingManualPriority = Boolean(data.paintingManualPriority);

    this.ldspStage = data.ldspStage;
    this.ldspStatus = data.ldspStatus;
    this.installationCorpusMdfComment = data.installationCorpusMdfComment || "";
    this.installationPaintedLdspComment = data.installationPaintedLdspComment || "";
    this.hasCorpusLdsp = data.hasCorpusLdsp !== false;
    this.hasCorpusMdf = Boolean(data.hasCorpusMdf);
    this.corpusMdfStage = data.corpusMdfStage || "";
    this.corpusMdfStatus = data.corpusMdfStatus || "";
    this.corpusMdfPaintingStage = data.corpusMdfPaintingStage || "";
    this.corpusMdfPaintingStatus = data.corpusMdfPaintingStatus || "";
    this.corpusMdfPaintingOperationCode = data.corpusMdfPaintingOperationCode || "";
    this.hasPaintedLdsp = Boolean(data.hasPaintedLdsp);
    this.paintedLdspStage = data.paintedLdspStage || "";
    this.paintedLdspStatus = data.paintedLdspStatus || "";
    this.paintedLdspPaintingStage = data.paintedLdspPaintingStage || "";
    this.paintedLdspPaintingStatus = data.paintedLdspPaintingStatus || "";
    this.paintedLdspPaintingOperationCode = data.paintedLdspPaintingOperationCode || "";
    this.mdfStage = data.mdfStage;
    this.mdfStatus = data.mdfStatus;
    this.hasMdfFacade = Boolean(data.hasMdfFacade);
    this.hasLdspFacade = Boolean(data.hasLdspFacade);
    this.veneerStage = data.veneerStage;
    this.veneerStatus = data.veneerStatus;
    this.hasVeneerFacade = Boolean(data.hasVeneerFacade);
    this.hasNoFacade = Boolean(data.hasNoFacade);
    this.facadePackagingStage = data.facadePackagingStage || "";
    this.facadePackagingStatus = data.facadePackagingStatus || "";
    this.mdfPaintingStage = data.mdfPaintingStage || "";
    this.mdfPaintingStatus = data.mdfPaintingStatus || "";
    this.mdfPaintingOperationCode = data.mdfPaintingOperationCode || "";
    this.veneerPaintingStage = data.veneerPaintingStage || "";
    this.veneerPaintingStatus = data.veneerPaintingStatus || "";
    this.veneerPaintingOperationCode = data.veneerPaintingOperationCode || "";
    this.mdfFacadePackagingStage = data.mdfFacadePackagingStage || "";
    this.mdfFacadePackagingStatus = data.mdfFacadePackagingStatus || "";
    this.veneerFacadePackagingStage = data.veneerFacadePackagingStage || "";
    this.veneerFacadePackagingStatus = data.veneerFacadePackagingStatus || "";

    this.sandingWorker = data.sandingWorker;

    this.sandingQueue = data.sandingQueue;

    this.whiteSandingReserveQueue = data.whiteSandingReserveQueue;

    this.sandingPlannedDate = data.sandingPlannedDate;

    this.painter = data.painter;

    this.primerQueue = data.primerQueue;

    this.paintingReserveQueue = data.paintingReserveQueue;

    this.paintingPlannedDate = data.paintingPlannedDate;

  }

  /**
   * Объект для передачи клиентскому приложению
   */
  toJSON() {

    return {

      key: this.key,

      manager: this.manager,

      contract: this.contract,

      client: this.client,

      product: this.product,

      displayName: this.displayName,

      stage: this.stage,

      responsible: this.responsible,

      constructor: this.constructor,

      queue: this.queue,

      startDate: this.startDate,

      plannedDate: this.plannedDate,

      deadline: this.deadline,

      deadlineMode: this.deadlineMode,

      manufacturingDays: this.manufacturingDays,

      manualDeadlineDate: this.manualDeadlineDate,

      manualPriority: this.manualPriority,

      actualStartDate: this.actualStartDate,

      actualStartHalf: this.actualStartHalf,

      remainingDuration: this.remainingDuration,

      duration: this.duration,

      tzDays: this.tzDays,

      proformaDays: this.proformaDays,

      kdDays: this.kdDays,

      kdReserveQueue: this.kdReserveQueue,

      proformaApprovedDate: this.proformaApprovedDate,

      finishDate: this.finishDate,

      comment: this.comment,

      installationCorpusComment: this.installationCorpusComment,
      installationFacadeComment: this.installationFacadeComment,
      installationMdfComment: this.installationMdfComment,
      installationVeneerComment: this.installationVeneerComment,
      installationLegacyComment: this.installationLegacyComment,
      claimComment: this.claimComment,
      manualClaim: this.manualClaim,

      status: this.status,

      cloudFolderUrl: this.cloudFolderUrl,

      productionDays: this.productionDays,

      lacquerOperations: this.lacquerOperations,

      routeExclusions: this.routeExclusions,

      paintingArea: this.paintingArea,

      backPaintingArea: this.backPaintingArea,

      frontPaintingArea: this.frontPaintingArea,

      polishingArea: this.polishingArea,

      ldspArea: this.ldspArea,
      mdfArea: this.mdfArea,
      lhdfArea: this.lhdfArea,
      veneerArea: this.veneerArea,

      facadeType: this.facadeType,

      paintingOperationCode: this.paintingOperationCode,

      productionQueue: this.productionQueue,

      paintingManualPriority: this.paintingManualPriority,

      ldspStage: this.ldspStage,
      ldspStatus: this.ldspStatus,
      installationCorpusMdfComment: this.installationCorpusMdfComment,
      installationPaintedLdspComment: this.installationPaintedLdspComment,
      hasCorpusLdsp: this.hasCorpusLdsp,
      hasCorpusMdf: this.hasCorpusMdf,
      corpusMdfStage: this.corpusMdfStage,
      corpusMdfStatus: this.corpusMdfStatus,
      corpusMdfPaintingStage: this.corpusMdfPaintingStage,
      corpusMdfPaintingStatus: this.corpusMdfPaintingStatus,
      corpusMdfPaintingOperationCode: this.corpusMdfPaintingOperationCode,
      hasPaintedLdsp: this.hasPaintedLdsp,
      paintedLdspStage: this.paintedLdspStage,
      paintedLdspStatus: this.paintedLdspStatus,
      paintedLdspPaintingStage: this.paintedLdspPaintingStage,
      paintedLdspPaintingStatus: this.paintedLdspPaintingStatus,
      paintedLdspPaintingOperationCode: this.paintedLdspPaintingOperationCode,
      mdfStage: this.mdfStage,
      mdfStatus: this.mdfStatus,
      hasMdfFacade: this.hasMdfFacade,
      hasLdspFacade: this.hasLdspFacade,
      veneerStage: this.veneerStage,
      veneerStatus: this.veneerStatus,
      hasVeneerFacade: this.hasVeneerFacade,
      hasNoFacade: this.hasNoFacade,
      facadePackagingStage: this.facadePackagingStage,
      facadePackagingStatus: this.facadePackagingStatus,
      mdfPaintingStage: this.mdfPaintingStage,
      mdfPaintingStatus: this.mdfPaintingStatus,
      mdfPaintingOperationCode: this.mdfPaintingOperationCode,
      veneerPaintingStage: this.veneerPaintingStage,
      veneerPaintingStatus: this.veneerPaintingStatus,
      veneerPaintingOperationCode: this.veneerPaintingOperationCode,
      mdfFacadePackagingStage: this.mdfFacadePackagingStage,
      mdfFacadePackagingStatus: this.mdfFacadePackagingStatus,
      veneerFacadePackagingStage: this.veneerFacadePackagingStage,
      veneerFacadePackagingStatus: this.veneerFacadePackagingStatus,

      sandingWorker: this.sandingWorker,

      sandingQueue: this.sandingQueue,

      whiteSandingReserveQueue: this.whiteSandingReserveQueue,

      sandingPlannedDate: this.sandingPlannedDate,

      painter: this.painter,

      primerQueue: this.primerQueue,

      paintingReserveQueue: this.paintingReserveQueue,

      paintingPlannedDate: this.paintingPlannedDate,

      weekday: ""

    };

  }

}

/**
 * Доменная модель "Изделие"
 */
class Order {

  constructor(data) {

    this.id = data.id;

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

    this.duration = data.duration;

    this.finishDate = data.finishDate;

    this.comment = data.comment;

  }

  /**
   * Объект для передачи клиентскому приложению
   */
  toJSON() {

    return {

      id: this.id,

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

      duration: this.duration,

      finishDate: this.finishDate,

      comment: this.comment,

      weekday: ""

    };

  }

}
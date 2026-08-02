class Api {

  static handleGet(e) {

    try {

      const resource = e.parameter.resource || "";

      switch (resource) {

        case "orders":
          return Response.success(
            OrderService.getOrders()
          );

        default:
          return Response.error("Unknown resource");

      }

    } catch (err) {

      return Response.error(err.message);

    }

  }

  static handlePost(e) {

    return Response.error("Not implemented");

  }

}
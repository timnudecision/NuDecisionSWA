define(["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    const httpTrigger = async function (context, req) {
        context.log("HTTP trigger function processed a request.");
        const name = req.query.name || (req.body && req.body.name);
        const responseMessage = name
            ? "Hello, " + name + ". This HTTP triggered function executed successfully."
            : "This HTTP triggered function executed successfully. Pass a name in the query string or in the request body for a personalized response.";
        context.res = {
            // status: 200, /* Defaults to 200 */
            body: responseMessage,
        };
    };
    exports.default = httpTrigger;
});
//# sourceMappingURL=index.js.map
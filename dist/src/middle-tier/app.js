/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file is the main Node.js server file that defines the express middleware.
 */
var __createBinding = (this && this.__createBinding) || (Object.create ? (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    var desc = Object.getOwnPropertyDescriptor(m, k);
    if (!desc || ("get" in desc ? !m.__esModule : desc.writable || desc.configurable)) {
      desc = { enumerable: true, get: function() { return m[k]; } };
    }
    Object.defineProperty(o, k2, desc);
}) : (function(o, m, k, k2) {
    if (k2 === undefined) k2 = k;
    o[k2] = m[k];
}));
var __setModuleDefault = (this && this.__setModuleDefault) || (Object.create ? (function(o, v) {
    Object.defineProperty(o, "default", { enumerable: true, value: v });
}) : function(o, v) {
    o["default"] = v;
});
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (k !== "default" && Object.prototype.hasOwnProperty.call(mod, k)) __createBinding(result, mod, k);
    __setModuleDefault(result, mod);
    return result;
};
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
define(["require", "exports", "http-errors", "path", "cookie-parser", "morgan", "express", "https", "office-addin-dev-certs", "./msgraph-helper", "./ssoauth-helper"], function (require, exports, createError, path, cookieParser, logger, express_1, https_1, office_addin_dev_certs_1, msgraph_helper_1, ssoauth_helper_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    createError = __importStar(createError);
    path = __importStar(path);
    cookieParser = __importStar(cookieParser);
    logger = __importStar(logger);
    express_1 = __importDefault(express_1);
    https_1 = __importDefault(https_1);
    if (process.env.NODE_ENV !== "production") {
        require("dotenv").config();
    }
    /* global console, process, require, __dirname */
    const app = (0, express_1.default)();
    const port = process.env.API_PORT || "3000";
    app.set("port", port);
    // view engine setup
    app.set("views", path.join(__dirname, "views"));
    app.set("view engine", "pug");
    app.use(logger("dev"));
    app.use(express_1.default.json());
    app.use(express_1.default.urlencoded({ extended: false }));
    app.use(cookieParser());
    /* Turn off caching when developing */
    if (process.env.NODE_ENV !== "production") {
        app.use(express_1.default.static(path.join(process.cwd(), "dist"), { etag: false }));
        app.use(function (req, res, next) {
            res.header("Cache-Control", "private, no-cache, no-store, must-revalidate");
            res.header("Expires", "-1");
            res.header("Pragma", "no-cache");
            next();
        });
    }
    else {
        // In production mode, let static files be cached.
        app.use(express_1.default.static(path.join(process.cwd(), "dist")));
    }
    const indexRouter = express_1.default.Router();
    indexRouter.get("/", function (req, res) {
        res.render("/taskpane.html");
    });
    app.use("/", indexRouter);
    // Middle-tier API calls
    // listen for 'ping' to verify service is running
    // Un comment for development debugging, but un needed for production deployment
    // app.get("/ping", function (req: any, res: any) {
    //   res.send(process.platform);
    // });
    app.get("/getuserdata", ssoauth_helper_1.validateJwt, msgraph_helper_1.getUserData);
    // Get the client side task pane files requested
    app.get("/taskpane.html", async (req, res) => {
        return res.sendfile("taskpane.html");
    });
    app.get("/fallbackauthdialog.html", async (req, res) => {
        return res.sendfile("fallbackauthdialog.html");
    });
    // Catch 404 and forward to error handler
    app.use(function (req, res, next) {
        next(createError(404));
    });
    // error handler
    app.use(function (err, req, res) {
        // set locals, only providing error in development
        res.locals.message = err.message;
        res.locals.error = req.app.get("env") === "development" ? err : {};
        // render the error page
        res.status(err.status || 500);
        res.render("error");
    });
    (0, office_addin_dev_certs_1.getHttpsServerOptions)().then((options) => {
        https_1.default
            .createServer(options, app)
            .listen(port, () => console.log(`Server running on ${port} in ${process.env.NODE_ENV} mode`));
    });
});
//# sourceMappingURL=app.js.map
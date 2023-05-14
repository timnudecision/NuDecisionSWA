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
define(["require", "exports", "https", "./ssoauth-helper", "http-errors"], function (require, exports, https, ssoauth_helper_1, createError) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.getGraphData = exports.getUserData = void 0;
    https = __importStar(https);
    createError = __importStar(createError);
    /* global process */
    const domain = "graph.microsoft.com";
    const version = "v1.0";
    async function getUserData(req, res, next) {
        const authorization = req.get("Authorization");
        await (0, ssoauth_helper_1.getAccessToken)(authorization)
            .then(async (graphTokenResponse) => {
            if (graphTokenResponse && (graphTokenResponse.claims || graphTokenResponse.error)) {
                res.send(graphTokenResponse);
            }
            else {
                const graphToken = graphTokenResponse.access_token;
                const graphUrlSegment = process.env.GRAPH_URL_SEGMENT || "/me";
                const graphQueryParamSegment = process.env.QUERY_PARAM_SEGMENT || "";
                const graphData = await getGraphData(graphToken, graphUrlSegment, graphQueryParamSegment);
                // If Microsoft Graph returns an error, such as invalid or expired token,
                // there will be a code property in the returned object set to a HTTP status (e.g. 401).
                // Relay it to the client. It will caught in the fail callback of `makeGraphApiCall`.
                if (graphData.code) {
                    next(createError(graphData.code, "Microsoft Graph error " + JSON.stringify(graphData)));
                }
                else {
                    res.send(graphData);
                }
            }
        })
            .catch((err) => {
            res.status(401).send(err.message);
            return;
        });
    }
    exports.getUserData = getUserData;
    async function getGraphData(accessToken, apiUrl, queryParams) {
        return new Promise((resolve, reject) => {
            const options = {
                host: domain,
                path: "/" + version + apiUrl + queryParams,
                method: "GET",
                headers: {
                    "Content-Type": "application/json",
                    Accept: "application/json",
                    Authorization: "Bearer " + accessToken,
                    "Cache-Control": "private, no-cache, no-store, must-revalidate",
                    Expires: "-1",
                    Pragma: "no-cache",
                },
            };
            https
                .get(options, (response) => {
                let body = "";
                response.on("data", (d) => {
                    body += d;
                });
                response.on("end", () => {
                    // The response from the OData endpoint might be an error, say a
                    // 401 if the endpoint requires an access token and it was invalid
                    // or expired. But a message is not an error in the call of https.get,
                    // so the "on('error', reject)" line below isn't triggered.
                    // So, the code distinguishes success (200) messages from error
                    // messages and sends a JSON object to the caller with either the
                    // requested OData or error information.
                    let error;
                    if (response.statusCode === 200) {
                        let parsedBody = JSON.parse(body);
                        resolve(parsedBody);
                    }
                    else {
                        error = new Error();
                        error.code = response.statusCode;
                        error.message = response.statusMessage;
                        // The error body sometimes includes an empty space
                        // before the first character, remove it or it causes an error.
                        body = body.trim();
                        error.bodyCode = JSON.parse(body).error.code;
                        error.bodyMessage = JSON.parse(body).error.message;
                        resolve(error);
                    }
                });
            })
                .on("error", reject);
        });
    }
    exports.getGraphData = getGraphData;
});
//# sourceMappingURL=msgraph-helper.js.map
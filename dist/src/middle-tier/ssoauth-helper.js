/*
 * Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. -->
 *
 * This file defines the routes within the authRoute router.
 */
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
define(["require", "exports", "node-fetch", "form-urlencoded", "jsonwebtoken", "jwks-rsa"], function (require, exports, node_fetch_1, form_urlencoded_1, jsonwebtoken_1, jwks_rsa_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.validateJwt = exports.getAccessToken = void 0;
    node_fetch_1 = __importDefault(node_fetch_1);
    form_urlencoded_1 = __importDefault(form_urlencoded_1);
    jsonwebtoken_1 = __importDefault(jsonwebtoken_1);
    /* global process, console */
    const DISCOVERY_KEYS_ENDPOINT = "https://login.microsoftonline.com/common/discovery/v2.0/keys";
    async function getAccessToken(authorization) {
        if (!authorization) {
            let error = new Error("No Authorization header was found.");
            return Promise.reject(error);
        }
        else {
            const scopeName = process.env.SCOPE || "User.Read";
            const [, /* schema */ assertion] = authorization.split(" ");
            const tokenScopes = jsonwebtoken_1.default.decode(assertion).scp.split(" ");
            const accessAsUserScope = tokenScopes.find((scope) => scope === "access_as_user");
            if (!accessAsUserScope) {
                throw new Error("Missing access_as_user");
            }
            const formParams = {
                client_id: process.env.CLIENT_ID,
                client_secret: process.env.CLIENT_SECRET,
                grant_type: "urn:ietf:params:oauth:grant-type:jwt-bearer",
                assertion: assertion,
                requested_token_use: "on_behalf_of",
                scope: [scopeName].join(" "),
            };
            const stsDomain = "https://login.microsoftonline.com";
            const tenant = "common";
            const tokenURLSegment = "oauth2/v2.0/token";
            const encodedForm = (0, form_urlencoded_1.default)(formParams);
            const tokenResponse = await (0, node_fetch_1.default)(`${stsDomain}/${tenant}/${tokenURLSegment}`, {
                method: "POST",
                body: encodedForm,
                headers: {
                    Accept: "application/json",
                    "Content-Type": "application/x-www-form-urlencoded",
                },
            });
            const json = await tokenResponse.json();
            return json;
        }
    }
    exports.getAccessToken = getAccessToken;
    function validateJwt(req, res, next) {
        const authHeader = req.headers.authorization;
        if (authHeader) {
            const token = authHeader.split(" ")[1];
            const validationOptions = {
                audience: process.env.CLIENT_ID,
            };
            jsonwebtoken_1.default.verify(token, getSigningKeys, validationOptions, (err) => {
                if (err) {
                    console.log(err);
                    return res.sendStatus(403);
                }
                next();
            });
        }
    }
    exports.validateJwt = validateJwt;
    function getSigningKeys(header, callback) {
        var client = new jwks_rsa_1.JwksClient({
            jwksUri: DISCOVERY_KEYS_ENDPOINT,
        });
        client.getSigningKey(header.kid, function (err, key) {
            callback(null, key.getPublicKey());
        });
    }
});
//# sourceMappingURL=ssoauth-helper.js.map
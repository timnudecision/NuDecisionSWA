// Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in the root of the repo.
/*
    This file provides the provides functionality to get Microsoft Graph data.
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
define(["require", "exports", "./message-helper", "jquery"], function (require, exports, message_helper_1, $) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.getUserData = void 0;
    $ = __importStar($);
    async function getUserData(middletierToken) {
        try {
            const response = await $.ajax({
                type: "GET",
                url: `/getuserdata`,
                headers: { Authorization: "Bearer " + middletierToken },
                cache: false,
            });
            return response;
        }
        catch (err) {
            (0, message_helper_1.showMessage)(`Error from middle tier. \n${err.responseText || err.message}`);
            throw err;
        }
    }
    exports.getUserData = getUserData;
});
//# sourceMappingURL=middle-tier-calls.js.map
/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
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
define(["require", "exports", "./../authentication/authhelper"], function (require, exports, fallbackAuthHelper) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.run1 = void 0;
    fallbackAuthHelper = __importStar(fallbackAuthHelper);
    Office.onReady((info) => {
        if (info.host === Office.HostType.Excel) {
            document.getElementById("getProfileButton").onclick = run1;
        }
    });
    async function run1() {
        InitAuth();
    }
    exports.run1 = run1;
    function InitAuth() {
        var authh = new fallbackAuthHelper.AuthHelper();
        authh.DialogFallback(setauthcookie);
    }
    async function setauthcookie(res, err) {
        localStorage.setItem("docregisterauth", res);
        console.log(res);
        const usrprofile = await getGraphdata(res);
        console.log(usrprofile);
        var roles = ["Admin", "User"]; // for local testing ONLY
        if (process.env.NODE_ENV !== "development")
            roles = await GetRoles(usrprofile.mail);
        roles = roles.concat(process.env.NODE_ENV);
        console.log(roles);
        document.getElementById("login-area").style.display = "none";
        document.getElementById("role-area").style.display = "block";
        document.getElementById("message-area").style.display = "block";
        if (roles.length == 0) {
            var li = document.createElement("li");
            li.append("No Roles defined for this login");
            document.getElementById("message-area").append(li);
        }
        else {
            for (var i = 0; i < roles.length; i++) {
                var li = document.createElement("li");
                li.append(roles[i]);
                document.getElementById("message-area").append(li);
            }
        }
    }
    async function GetRoles(email) {
        const url = "https://" + window.location.host + "/api/Auth?name=" + email;
        let res = await fetch(url);
        let resbody = await res.json();
        return resbody;
        //console.log(resbody);
    }
    async function getGraphdata(accesstoken) {
        const domain = "https://graph.microsoft.com/v1.0/me";
        const headers1 = {
            "Content-Type": "application/json",
            Accept: "application/json",
            Authorization: "Bearer " + accesstoken,
        };
        let res = await fetch(domain, { headers: headers1 });
        let resbody = await res.json();
        return resbody;
    }
});
//# sourceMappingURL=taskpane.js.map
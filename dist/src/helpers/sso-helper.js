/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
define(["require", "exports", "./fallbackauthdialog.js", "./middle-tier-calls", "./message-helper", "./error-handler"], function (require, exports, fallbackauthdialog_js_1, middle_tier_calls_1, message_helper_1, error_handler_1) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.ShowDial = exports.getUserProfile = void 0;
    /* global Office */
    let retryGetMiddletierToken = 0;
    async function getUserProfile(callback) {
        try {
            let middletierToken = await Office.auth.getAccessToken({
                allowSignInPrompt: true,
                allowConsentPrompt: true,
                forMSGraphAccess: true,
            });
            console.log('getting userdata3');
            let response = await (0, middle_tier_calls_1.getUserData)(middletierToken);
            if (!response) {
                throw new Error("Middle tier didn't respond");
            }
            else if (response.claims) {
                // Microsoft Graph requires an additional form of authentication. Have the Office host
                // get a new token using the Claims string, which tells AAD to prompt the user for all
                // required forms of authentication.
                let mfaMiddletierToken = await Office.auth.getAccessToken({
                    authChallenge: response.claims,
                });
                console.log('getting userdata1');
                response = (0, middle_tier_calls_1.getUserData)(mfaMiddletierToken);
            }
            // AAD errors are returned to the client with HTTP code 200, so they do not trigger
            // the catch block below.
            if (response.error) {
                handleAADErrors(response, callback);
            }
            else {
                callback(response);
            }
        }
        catch (exception) {
            // if handleClientSideErrors returns true then we will try to authenticate via the fallback
            // dialog rather than simply throw and error
            if (exception.code) {
                if ((0, error_handler_1.handleClientSideErrors)(exception)) {
                    (0, fallbackauthdialog_js_1.dialogFallback)(callback);
                }
            }
            else {
                (0, message_helper_1.showMessage)("EXCEPTION: " + JSON.stringify(exception));
                throw exception;
            }
        }
    }
    exports.getUserProfile = getUserProfile;
    function handleAADErrors(response, callback) {
        // On rare occasions the middle tier token is unexpired when Office validates it,
        // but expires by the time it is sent to AAD for exchange. AAD will respond
        // with "The provided value for the 'assertion' is not valid. The assertion has expired."
        // Retry the call of getAccessToken (no more than once). This time Office will return a
        // new unexpired middle tier token.
        if (response.error_description.indexOf("AADSTS500133") !== -1 && retryGetMiddletierToken <= 0) {
            retryGetMiddletierToken++;
            getUserProfile(callback);
        }
        else {
            (0, fallbackauthdialog_js_1.dialogFallback2)(callback);
        }
    }
    async function ShowDial(callback) {
        (0, fallbackauthdialog_js_1.dialogFallback2)(callback);
    }
    exports.ShowDial = ShowDial;
});
//# sourceMappingURL=sso-helper.js.map
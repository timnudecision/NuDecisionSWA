/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
define(["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.hideMessage = exports.clearMessage = exports.showMessage = void 0;
    /* global document */
    function showMessage(text) {
        document.getElementById("message-area").style.display = "flex";
        document.getElementById("message-area").innerText = text;
    }
    exports.showMessage = showMessage;
    function clearMessage() {
        document.getElementById("message-area").style.display = "flex";
        document.getElementById("message-area").innerText = "---<br>";
    }
    exports.clearMessage = clearMessage;
    function hideMessage() {
        document.getElementById("message-area").style.display = "none";
        document.getElementById("message-area").innerText = "---<br>";
    }
    exports.hideMessage = hideMessage;
});
//# sourceMappingURL=message-helper.js.map
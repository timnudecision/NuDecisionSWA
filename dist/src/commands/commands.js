/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
define(["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.action = void 0;
    /* global global, Office, self, window */
    Office.onReady(() => {
        // If needed, Office.js is ready to be called
    });
    /**
     * Shows a notification when the add-in command is executed.
     * @param event
     */
    function action(event) {
        const message = {
            type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
            message: "Performed action.",
            icon: "Icon.80x80",
            persistent: true,
        };
        // Show a notification message
        Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
        // Be sure to indicate when the add-in command function is complete
        event.completed();
    }
    exports.action = action;
    function getGlobal() {
        return typeof self !== "undefined"
            ? self
            : typeof window !== "undefined"
                ? window
                : typeof global !== "undefined"
                    ? global
                    : undefined;
    }
    const g = getGlobal();
    // the add-in command functions need to be available in global scope
    g.action;
});
//# sourceMappingURL=commands.js.map
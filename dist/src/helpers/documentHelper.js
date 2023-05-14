/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */
define(["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.filterUserProfileInfo = void 0;
    function filterUserProfileInfo(result) {
        let userProfileInfo = [];
        userProfileInfo.push(result["displayName"]);
        userProfileInfo.push(result["jobTitle"]);
        userProfileInfo.push(result["mail"]);
        userProfileInfo.push(result["mobilePhone"]);
        userProfileInfo.push(result["officeLocation"]);
        return userProfileInfo;
    }
    exports.filterUserProfileInfo = filterUserProfileInfo;
});
//# sourceMappingURL=documentHelper.js.map
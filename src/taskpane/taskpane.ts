/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

import * as fallbackAuthHelper from "./../authentication/authhelper";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("getProfileButton").onclick = run1;
  }
});

export async function run1() {
  InitAuth();
}

function InitAuth() {
  var authh = new fallbackAuthHelper.AuthHelper();
  authh.DialogFallback(setauthcookie);
}

async function setauthcookie(res, err) {
  localStorage.setItem("docregisterauth", res);

  console.log(JSON.stringify(res));
  const usrprofile = await getGraphdata(res);
  console.log(JSON.stringify(usrprofile));

  var roles = ["Admin", "User"]; // for local testing ONLY
  if (process.env.NODE_ENV !== "development") roles = await GetRoles(usrprofile.mail);
  roles = roles.concat(process.env.NODE_ENV);
  console.log(JSON.stringify(roles));

  document.getElementById("login-area").style.display = "none";
  document.getElementById("role-area").style.display = "block";
  document.getElementById("message-area").style.display = "block";

  var li = document.createElement("li");
  const txt =
    roles.length == 0
      ? "No Roles defined for this login"
      : `Roles for ${usrprofile.displayName}, ${usrprofile.mail} are:\n` + roles.join(", ");
  li.append(txt);
  document.getElementById("message-area").append(li);
}

async function GetRoles(email) {
  const url = "https://" + window.location.host + "/api/Auth?name=" + email;
  let res = await fetch(url);
  let resbody = await res.json();
  return resbody;
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

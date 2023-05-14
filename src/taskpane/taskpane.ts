/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

//const fallbackAuthHelper = require("./../authentication/authhelper");
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
  console.log(res);
  const usrprofile = await getGraphdata(res);
  console.log(usrprofile);
  var roles = ["Admin", "User"]; // for local testing ONLY
  if (process.env.NODE_ENV !== "development") roles = await GetRoles(usrprofile.mail);
  roles = roles.concat(process.env.NODE_ENV);
  console.log(roles);
  document.getElementById("login-area").style.display = "none";
  document.getElementById("role-area").style.display = "block";
  document.getElementById("message-area").style.display = "block";
  if (roles.length == 0) {
    var li = document.createElement("li");
    li.append("No Roles defined for this login");
    document.getElementById("message-area").append(li);
  } else {
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

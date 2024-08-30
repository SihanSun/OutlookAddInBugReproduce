/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import  * as msal from "@azure/msal-browser";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = signIn;
  }
});

export async function run() {
  /**
   * Insert your Outlook code here
   */

  const item = Office.context.mailbox.item;
  let insertAt = document.getElementById("item-subject");
  let label = document.createElement("b").appendChild(document.createTextNode("Subject: "));
  insertAt.appendChild(label);
  insertAt.appendChild(document.createElement("br"));
  insertAt.appendChild(document.createTextNode(item.subject));
  insertAt.appendChild(document.createElement("br"));
}


const msalConfig = {
  auth: {
      clientId: 'aea232ad-42b1-46a3-bf4c-e7d55883d789', // This is the ONLY mandatory field that you need to supply.
      authority: 'https://login.microsoftonline.com/common', // Replace the placeholder with your tenant subdomain        
      // navigateToLoginRequestUrl: true, // If "true", will navigate back to the original request location before processing the auth code response.
  },
  cache: {
      cacheLocation: 'sessionStorage', // Configures cache location. "sessionStorage" is more secure, but "localStorage" gives you SSO.
      storeAuthStateInCookie: false, // set this to true if you have to support IE
  },
  system: {
      loggerOptions: {
          loggerCallback: (level, message, containsPii) => {
              if (containsPii) {
                  return;
              }
              switch (level) {
                  case msal.LogLevel.Error:
                      console.error(message);
                      return;
                  case msal.LogLevel.Info:
                      console.info(message);
                      return;
                  case msal.LogLevel.Verbose:
                      console.debug(message);
                      return;
                  case msal.LogLevel.Warning:
                      console.warn(message);
                      return;
              }
          },
      },
  },
};

const loginRequest = {
  scopes: ["User.Read"]
};

let pca = undefined;

console.log("creating pca");
msal.createNestablePublicClientApplication(msalConfig).then((result) => {
  console.log(result);
  pca = result;
  console.log(pca);
}).catch(error => {
  console.log(error);
});

export function signIn() {
  pca.loginPopup(loginRequest).then(function(response) {
    console
    console.log(response);
  }).catch((error) => {
    console.log(error);
  })
}



/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */

import  * as msal from "@azure/msal-browser";

const setResult = (result) => {
  const labelForResult = document.getElementById("result");
  if (labelForResult) {
    labelForResult.textContent = result;
  }
};

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("mailboxEmail").textContent = getMailboxEmail();
    testAuthContext();
    initialize();
  }
});

// fccd3bcf-08f0-4b8a-b36f-520cfaa4ab51
// aea232ad-42b1-46a3-bf4c-e7d55883d789

const msalConfig = {
  auth: {
      clientId: 'fccd3bcf-08f0-4b8a-b36f-520cfaa4ab51', // This is the ONLY mandatory field that you need to supply.
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

const getMailboxEmail = () => {
  return Office.context.mailbox.userProfile.emailAddress;
};

const loginRequest = {
  scopes: ["User.Read"]
};
let pca = undefined;

const initialize = () => {
  msal.createNestablePublicClientApplication(msalConfig).then((result) => {
    pca = result;
    
    const activeAccount = pca.getActiveAccount();
    if (activeAccount) {
      document.getElementById("graphUsername").textContent = activeAccount.username;
      hideSignin();
    } else {
      showSignin();
    }
  }).catch(error => {
    console.log(error);
  });
}

function signIn() {
  pca.loginPopup(loginRequest).then(function(response) {
    pca.setActiveAccount(response.account);
    document.getElementById("graphUsername").textContent = pca.getActiveAccount().username;
    hideSignin();
  }).catch((error) => {
    console.error(error);
  })
}

function showSignin() {
  const signInButton = document.getElementById("signin");
  signInButton.style.display = "block";
  signInButton.onclick = signIn;
}

function hideSignin() {
  const signInButton = document.getElementById("signin");
  signInButton.style.display = "none";
  signInButton.onclick = null;
}

function testAuthContext() {
  Office.auth.getAuthContext().then((authContext) => {
    document.getElementById('authContext').textContent = "Available";
  }).catch(() => {
    document.getElementById('authContext').textContent = "Not Available";
  });
}



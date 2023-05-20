//import * as msal from '@azure/msal-browser';
//require ("https://alcdn.msauth.net/browser/2.13.1/js/msal-browser.min.js");

// /import { logger } from "./Shared/Helper";

$(function(){

    logger.log('Auth now');
    
    //logincommon.auth();
    Office.onReady().then(function(info)
    {
        console.log('office context:',Office.context.ui);
        console.log(info);
        logger.log(info);
        logincommon.auth();
    });

})

var loginfunctions={};

loginfunctions.loginfn=()=>
{
    $('#contentdiv').append("<h2>ok</h2>");
    $('#contentdiv').append("<h2>office ready</h2>");
    logincommon.login();
  
}


loginfunctions.logoutfn=()=>
{
    Office.onReady().then(function()
    {
         logincommon.logout();

    });
}

var logincommon=
{
     
    msalConfig:{
        auth: {           
            clientId: '121964d0-6350-438e-b38a-54bfedab82bf',  // Azure application ID         
            authority: 'https://login.microsoftonline.com/common',
          },
          cache: {
            cacheLocation: "localStorage", // Needed to avoid "User login is required" error.
            storeAuthStateInCookie: true // Recommended to avoid certain IE/Edge issues.
          }
    },

    _msalInstance:null,
    get msalInstance()
    {
     if(!logincommon._msalInstance) logincommon._msalInstance=new msal.PublicClientApplication(logincommon.msalConfig);
     return logincommon._msalInstance;
    }



};

logincommon.msalredirecthander=function(res){
    
    console.log(res);
    const currentAccounts = logincommon.msalInstance.getAllAccounts();
    if (currentAccounts === null) {
        logger.log( "no accounts detected");
    }
    else if (currentAccounts.length > 1) 
    {
        // Add choose account code here
        logger.log("multiple accounts");
    } else if (currentAccounts.length === 1) {
        username = currentAccounts[0].username;
        logger.log("1 accounts: "+username);
    }
   
    if(res)
    {
        authCallback(null,res);
    }
}

logincommon.login=function()
{
    try
    {
        $('#contentdiv').append("<h2>Tryign login</h2>");
        //if(!logincommon.msalInstance) logincommon.msalInstance=new msal.PublicClientApplication(logincommon.msalConfig);
        var loginRequest = {
            scopes: ["User.Read","Sites.ReadWrite.All",
             //"https://graph.microsoft.com/Sites.ReadWrite.All",
             //"https://incitias.sharepoint.com/AllSites.Manage",
             //"https://incitias.sharepoint.com/User.Read.All"
             ] // optional Array<string>
        };
        logincommon.msalInstance.handleRedirectPromise().then(
            logincommon.msalredirecthander,
            function(err){logger.log('logn redirect error: '+JSON.stringify(err));}
        );
        logincommon.msalInstance.loginRedirect(loginRequest);
    }
    catch(loginerr)
    {
        $('#contentdiv').append("<h2>Login error: "+JSON.stringify(loginerr)+"</h2>");

        logger.log('Login Error: '+JSON.stringify(loginerr));
    }


}

logincommon.logout=function()
{
    try
    {
        //if(!logincommon.msalInstance) logincommon.msalInstance=new msal.PublicClientApplication(logincommon.msalConfig);

        const logoutRequest = {
            account: logincommon.msalInstance.getAccountByUsername("admin@sha2013.onmicrosoft.com")
        }
        
        logincommon.msalInstance.logout(logoutRequest);
    }
    catch(loginerr)
    {
        logger.log('Login Error: '+loginerr);
    }
}

logincommon.auth=async function()
{
    try
    {
        logger.log('Auth now in function');

        var currentAccount;
        var allacocunts=logincommon.msalInstance.getAllAccounts();//getAccountByUsername("admin@sha2013.onmicrosoft.com");
        //var currentAccount = logincommon.msalInstance.getAccountByUsername("ankur.agarwal@incitias.com");
        //logger.log('current account',currentAccount.username);
        if(allacocunts && allacocunts.length>0) currentAccount=allacocunts[0];
        else currentAccount=null;
        console.log(currentAccount);
        var asyncloginRequest = {
         scopes: [
             "User.Read","Sites.ReadWrite.All",
        // "https://incitias.sharepoint.com/AllSites.Manage",
        // "https://incitias.sharepoint.com/User.Read.All"
        ], // optional Array<string>
         account: currentAccount,
         //prompt:'consent',
         forceRefresh: false,
        };
 
        var loginRequest = 
        {
         scopes: [
             "User.Read","Sites.ReadWrite.All",
            // "https://incitias.sharepoint.com/AllSites.Manage",
         //"https://incitias.sharepoint.com/User.Read.All"
        ],
         // prompt:'consent',

         //loginHint: currentAccount.username // optional Array<string>
        };
        logincommon.msalInstance.handleRedirectPromise().then(
            logincommon.msalredirecthander,
            function(err){logger.log('logn redirect error: '+JSON.stringify(err));}
        );
        var tksil=await logincommon.msalInstance.acquireTokenSilent(asyncloginRequest).then(function(tokenResponse)
         {
         // Do something with the tokenResponse
           logger.log('got it '+ JSON.stringify(tokenResponse));
           console.log('got it',tokenResponse);
           return tokenResponse;
         }).catch(function(error) 
         {
             logger.log('err: ',error);
 
             // if (error instanceof InteractionRequiredAuthError)
             //  {
                 // fallback to interaction when silent call fails
                 return logincommon.msalInstance.acquireTokenRedirect(loginRequest);
              //}
         });
 
         logger.log(JSON.stringify(tksil));
         console.log('now auth call back',tksil);
         authCallback(null,tksil);
    }
    catch(autherr)
    {
        console.log('auth error',autherr);

        logger.log("Error in Auth");
        logger.log("",autherr);
    }
}


function authCallback(error, response) {
    console.log('in authcallbacl',response.tokenType);
    if (error) {
      logger.log(error);
      //Office.context.ui.messageParent(JSON.stringify({ status: "failure", result: error }));
    } else {
      if (response.tokenType === "id_token") {
        logger.log("id_token"+response.idToken.rawIdToken);
        //localStorage.setItem("loggedIn", "yes");
      } else {
        logger.log("token type is:" + response.tokenType);
        console.log('token type is',response.tokenType);
     
        Office.onReady().then(info=>
        {
            Office.context.ui.messageParent(JSON.stringify({ status: "success", result: response.accessToken }));

        });
      }
    }
  }


var logger={};

logger.log=async (logmsg,logobj)=>
{  
    try
    {
    //console.log(logmsg,logobj);    
    if(logobj)logmsg=logmsg+": "+JSON.stringify(logobj);
    // if (Object.prototype.toString.call(logmsg) === '[object Object]')
    // logmsg=JSON.stringify(logmsg);
    
    console.log(logmsg);
    

    }
    catch(err){ console.log('eeror in logger',err); }

}



define(["require", "exports"], function (require, exports) {
    "use strict";
    Object.defineProperty(exports, "__esModule", { value: true });
    exports.AuthHelper = void 0;
    function AuthHelper() {
        var loginDialog;
        var parentcallback;
        this.DialogFallback = async (_pcallback) => {
            parentcallback = _pcallback;
            // We fall back to Dialog API for any error.
            const url = "/login.html";
            // Office.onReady().then(r=>{
            ShowLoginPopup(url);
            // })
            //const tk=await OfficeRuntime.auth.getAccessToken();
            //GetGraphData(tk);
        };
        var ShowLoginPopup = (url) => {
            var fullUrl = window.location.protocol +
                "//" +
                window.location.hostname +
                (window.location.port ? ":" + window.location.port : "") +
                url;
            var _this = this;
            Office.context.ui.displayDialogAsync(fullUrl, { height: 60, width: 30 }, function (result) {
                loginDialog = result.value;
                loginDialog.addEventHandler(Office.EventType.DialogMessageReceived, ProcessMessage);
            });
        };
        var ProcessMessage = async (arg) => {
            console.log("proc message");
            logger.log("Message received in processMessage: " + JSON.stringify(arg));
            let messageFromDialog = JSON.parse(arg.message);
            loginDialog.close();
            if (messageFromDialog.status === "success") {
                // We now have a valid access token.
                try {
                    logger.log("final call back: " + messageFromDialog.result);
                    //GetGraphData(messageFromDialog.result);
                    if (parentcallback) {
                        logger.log("got parent call back");
                        parentcallback(messageFromDialog.result, null);
                    }
                    else
                        logger.log("parent call back null");
                }
                catch (ex1) {
                    logger.log(ex1);
                    if (parentcallback)
                        parentcallback(null, ex1);
                    else
                        logger.log("parent call back null ex1");
                }
            }
            else {
                //this.loginDialog.close();
                if (parentcallback)
                    parentcallback(null, messageFromDialog);
                else
                    logger.log("parent call back failure");
            }
        };
    }
    exports.AuthHelper = AuthHelper;
    var logger = {};
    logger.log = async (logmsg, logobj) => {
        try {
            //console.log(logmsg,logobj);
            if (logobj)
                logmsg = logmsg + ": " + JSON.stringify(logobj);
            // if (Object.prototype.toString.call(logmsg) === '[object Object]')
            // logmsg=JSON.stringify(logmsg);
        }
        catch (err) {
            console.log("eeror in logger", err);
        }
    };
});
//# sourceMappingURL=authhelper.js.map
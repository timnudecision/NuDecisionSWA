export function AuthHelper() {
  var loginDialog;
  var parentcallback;

  this.DialogFallback = async (_pcallback) => {
    parentcallback = _pcallback;

    // We fall back to Dialog API for any error.
    const url = "/login.html";
    ShowLoginPopup(url);
  };

  var ShowLoginPopup = (url) => {
    var fullUrl =
      window.location.protocol +
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
    let messageFromDialog = JSON.parse(arg.message);
    loginDialog.close();

    if (messageFromDialog.status === "success") {
      // We now have a valid access token.
      try {
        if (parentcallback) parentcallback(messageFromDialog.result, null);
      } catch (ex1) {
        if (parentcallback) parentcallback(null, ex1);
      }
    } else {
      if (parentcallback) parentcallback(null, messageFromDialog);
    }
  };
}
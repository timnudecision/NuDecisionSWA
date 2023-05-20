/******/ (function() { // webpackBootstrap
var __webpack_exports__ = {};
/*!*************************************!*\
  !*** ./src/authentication/login.js ***!
  \*************************************/
function _typeof(obj) { "@babel/helpers - typeof"; return _typeof = "function" == typeof Symbol && "symbol" == typeof Symbol.iterator ? function (obj) { return typeof obj; } : function (obj) { return obj && "function" == typeof Symbol && obj.constructor === Symbol && obj !== Symbol.prototype ? "symbol" : typeof obj; }, _typeof(obj); }
function _regeneratorRuntime() { "use strict"; /*! regenerator-runtime -- Copyright (c) 2014-present, Facebook, Inc. -- license (MIT): https://github.com/facebook/regenerator/blob/main/LICENSE */ _regeneratorRuntime = function _regeneratorRuntime() { return exports; }; var exports = {}, Op = Object.prototype, hasOwn = Op.hasOwnProperty, defineProperty = Object.defineProperty || function (obj, key, desc) { obj[key] = desc.value; }, $Symbol = "function" == typeof Symbol ? Symbol : {}, iteratorSymbol = $Symbol.iterator || "@@iterator", asyncIteratorSymbol = $Symbol.asyncIterator || "@@asyncIterator", toStringTagSymbol = $Symbol.toStringTag || "@@toStringTag"; function define(obj, key, value) { return Object.defineProperty(obj, key, { value: value, enumerable: !0, configurable: !0, writable: !0 }), obj[key]; } try { define({}, ""); } catch (err) { define = function define(obj, key, value) { return obj[key] = value; }; } function wrap(innerFn, outerFn, self, tryLocsList) { var protoGenerator = outerFn && outerFn.prototype instanceof Generator ? outerFn : Generator, generator = Object.create(protoGenerator.prototype), context = new Context(tryLocsList || []); return defineProperty(generator, "_invoke", { value: makeInvokeMethod(innerFn, self, context) }), generator; } function tryCatch(fn, obj, arg) { try { return { type: "normal", arg: fn.call(obj, arg) }; } catch (err) { return { type: "throw", arg: err }; } } exports.wrap = wrap; var ContinueSentinel = {}; function Generator() {} function GeneratorFunction() {} function GeneratorFunctionPrototype() {} var IteratorPrototype = {}; define(IteratorPrototype, iteratorSymbol, function () { return this; }); var getProto = Object.getPrototypeOf, NativeIteratorPrototype = getProto && getProto(getProto(values([]))); NativeIteratorPrototype && NativeIteratorPrototype !== Op && hasOwn.call(NativeIteratorPrototype, iteratorSymbol) && (IteratorPrototype = NativeIteratorPrototype); var Gp = GeneratorFunctionPrototype.prototype = Generator.prototype = Object.create(IteratorPrototype); function defineIteratorMethods(prototype) { ["next", "throw", "return"].forEach(function (method) { define(prototype, method, function (arg) { return this._invoke(method, arg); }); }); } function AsyncIterator(generator, PromiseImpl) { function invoke(method, arg, resolve, reject) { var record = tryCatch(generator[method], generator, arg); if ("throw" !== record.type) { var result = record.arg, value = result.value; return value && "object" == _typeof(value) && hasOwn.call(value, "__await") ? PromiseImpl.resolve(value.__await).then(function (value) { invoke("next", value, resolve, reject); }, function (err) { invoke("throw", err, resolve, reject); }) : PromiseImpl.resolve(value).then(function (unwrapped) { result.value = unwrapped, resolve(result); }, function (error) { return invoke("throw", error, resolve, reject); }); } reject(record.arg); } var previousPromise; defineProperty(this, "_invoke", { value: function value(method, arg) { function callInvokeWithMethodAndArg() { return new PromiseImpl(function (resolve, reject) { invoke(method, arg, resolve, reject); }); } return previousPromise = previousPromise ? previousPromise.then(callInvokeWithMethodAndArg, callInvokeWithMethodAndArg) : callInvokeWithMethodAndArg(); } }); } function makeInvokeMethod(innerFn, self, context) { var state = "suspendedStart"; return function (method, arg) { if ("executing" === state) throw new Error("Generator is already running"); if ("completed" === state) { if ("throw" === method) throw arg; return doneResult(); } for (context.method = method, context.arg = arg;;) { var delegate = context.delegate; if (delegate) { var delegateResult = maybeInvokeDelegate(delegate, context); if (delegateResult) { if (delegateResult === ContinueSentinel) continue; return delegateResult; } } if ("next" === context.method) context.sent = context._sent = context.arg;else if ("throw" === context.method) { if ("suspendedStart" === state) throw state = "completed", context.arg; context.dispatchException(context.arg); } else "return" === context.method && context.abrupt("return", context.arg); state = "executing"; var record = tryCatch(innerFn, self, context); if ("normal" === record.type) { if (state = context.done ? "completed" : "suspendedYield", record.arg === ContinueSentinel) continue; return { value: record.arg, done: context.done }; } "throw" === record.type && (state = "completed", context.method = "throw", context.arg = record.arg); } }; } function maybeInvokeDelegate(delegate, context) { var methodName = context.method, method = delegate.iterator[methodName]; if (undefined === method) return context.delegate = null, "throw" === methodName && delegate.iterator.return && (context.method = "return", context.arg = undefined, maybeInvokeDelegate(delegate, context), "throw" === context.method) || "return" !== methodName && (context.method = "throw", context.arg = new TypeError("The iterator does not provide a '" + methodName + "' method")), ContinueSentinel; var record = tryCatch(method, delegate.iterator, context.arg); if ("throw" === record.type) return context.method = "throw", context.arg = record.arg, context.delegate = null, ContinueSentinel; var info = record.arg; return info ? info.done ? (context[delegate.resultName] = info.value, context.next = delegate.nextLoc, "return" !== context.method && (context.method = "next", context.arg = undefined), context.delegate = null, ContinueSentinel) : info : (context.method = "throw", context.arg = new TypeError("iterator result is not an object"), context.delegate = null, ContinueSentinel); } function pushTryEntry(locs) { var entry = { tryLoc: locs[0] }; 1 in locs && (entry.catchLoc = locs[1]), 2 in locs && (entry.finallyLoc = locs[2], entry.afterLoc = locs[3]), this.tryEntries.push(entry); } function resetTryEntry(entry) { var record = entry.completion || {}; record.type = "normal", delete record.arg, entry.completion = record; } function Context(tryLocsList) { this.tryEntries = [{ tryLoc: "root" }], tryLocsList.forEach(pushTryEntry, this), this.reset(!0); } function values(iterable) { if (iterable) { var iteratorMethod = iterable[iteratorSymbol]; if (iteratorMethod) return iteratorMethod.call(iterable); if ("function" == typeof iterable.next) return iterable; if (!isNaN(iterable.length)) { var i = -1, next = function next() { for (; ++i < iterable.length;) if (hasOwn.call(iterable, i)) return next.value = iterable[i], next.done = !1, next; return next.value = undefined, next.done = !0, next; }; return next.next = next; } } return { next: doneResult }; } function doneResult() { return { value: undefined, done: !0 }; } return GeneratorFunction.prototype = GeneratorFunctionPrototype, defineProperty(Gp, "constructor", { value: GeneratorFunctionPrototype, configurable: !0 }), defineProperty(GeneratorFunctionPrototype, "constructor", { value: GeneratorFunction, configurable: !0 }), GeneratorFunction.displayName = define(GeneratorFunctionPrototype, toStringTagSymbol, "GeneratorFunction"), exports.isGeneratorFunction = function (genFun) { var ctor = "function" == typeof genFun && genFun.constructor; return !!ctor && (ctor === GeneratorFunction || "GeneratorFunction" === (ctor.displayName || ctor.name)); }, exports.mark = function (genFun) { return Object.setPrototypeOf ? Object.setPrototypeOf(genFun, GeneratorFunctionPrototype) : (genFun.__proto__ = GeneratorFunctionPrototype, define(genFun, toStringTagSymbol, "GeneratorFunction")), genFun.prototype = Object.create(Gp), genFun; }, exports.awrap = function (arg) { return { __await: arg }; }, defineIteratorMethods(AsyncIterator.prototype), define(AsyncIterator.prototype, asyncIteratorSymbol, function () { return this; }), exports.AsyncIterator = AsyncIterator, exports.async = function (innerFn, outerFn, self, tryLocsList, PromiseImpl) { void 0 === PromiseImpl && (PromiseImpl = Promise); var iter = new AsyncIterator(wrap(innerFn, outerFn, self, tryLocsList), PromiseImpl); return exports.isGeneratorFunction(outerFn) ? iter : iter.next().then(function (result) { return result.done ? result.value : iter.next(); }); }, defineIteratorMethods(Gp), define(Gp, toStringTagSymbol, "Generator"), define(Gp, iteratorSymbol, function () { return this; }), define(Gp, "toString", function () { return "[object Generator]"; }), exports.keys = function (val) { var object = Object(val), keys = []; for (var key in object) keys.push(key); return keys.reverse(), function next() { for (; keys.length;) { var key = keys.pop(); if (key in object) return next.value = key, next.done = !1, next; } return next.done = !0, next; }; }, exports.values = values, Context.prototype = { constructor: Context, reset: function reset(skipTempReset) { if (this.prev = 0, this.next = 0, this.sent = this._sent = undefined, this.done = !1, this.delegate = null, this.method = "next", this.arg = undefined, this.tryEntries.forEach(resetTryEntry), !skipTempReset) for (var name in this) "t" === name.charAt(0) && hasOwn.call(this, name) && !isNaN(+name.slice(1)) && (this[name] = undefined); }, stop: function stop() { this.done = !0; var rootRecord = this.tryEntries[0].completion; if ("throw" === rootRecord.type) throw rootRecord.arg; return this.rval; }, dispatchException: function dispatchException(exception) { if (this.done) throw exception; var context = this; function handle(loc, caught) { return record.type = "throw", record.arg = exception, context.next = loc, caught && (context.method = "next", context.arg = undefined), !!caught; } for (var i = this.tryEntries.length - 1; i >= 0; --i) { var entry = this.tryEntries[i], record = entry.completion; if ("root" === entry.tryLoc) return handle("end"); if (entry.tryLoc <= this.prev) { var hasCatch = hasOwn.call(entry, "catchLoc"), hasFinally = hasOwn.call(entry, "finallyLoc"); if (hasCatch && hasFinally) { if (this.prev < entry.catchLoc) return handle(entry.catchLoc, !0); if (this.prev < entry.finallyLoc) return handle(entry.finallyLoc); } else if (hasCatch) { if (this.prev < entry.catchLoc) return handle(entry.catchLoc, !0); } else { if (!hasFinally) throw new Error("try statement without catch or finally"); if (this.prev < entry.finallyLoc) return handle(entry.finallyLoc); } } } }, abrupt: function abrupt(type, arg) { for (var i = this.tryEntries.length - 1; i >= 0; --i) { var entry = this.tryEntries[i]; if (entry.tryLoc <= this.prev && hasOwn.call(entry, "finallyLoc") && this.prev < entry.finallyLoc) { var finallyEntry = entry; break; } } finallyEntry && ("break" === type || "continue" === type) && finallyEntry.tryLoc <= arg && arg <= finallyEntry.finallyLoc && (finallyEntry = null); var record = finallyEntry ? finallyEntry.completion : {}; return record.type = type, record.arg = arg, finallyEntry ? (this.method = "next", this.next = finallyEntry.finallyLoc, ContinueSentinel) : this.complete(record); }, complete: function complete(record, afterLoc) { if ("throw" === record.type) throw record.arg; return "break" === record.type || "continue" === record.type ? this.next = record.arg : "return" === record.type ? (this.rval = this.arg = record.arg, this.method = "return", this.next = "end") : "normal" === record.type && afterLoc && (this.next = afterLoc), ContinueSentinel; }, finish: function finish(finallyLoc) { for (var i = this.tryEntries.length - 1; i >= 0; --i) { var entry = this.tryEntries[i]; if (entry.finallyLoc === finallyLoc) return this.complete(entry.completion, entry.afterLoc), resetTryEntry(entry), ContinueSentinel; } }, catch: function _catch(tryLoc) { for (var i = this.tryEntries.length - 1; i >= 0; --i) { var entry = this.tryEntries[i]; if (entry.tryLoc === tryLoc) { var record = entry.completion; if ("throw" === record.type) { var thrown = record.arg; resetTryEntry(entry); } return thrown; } } throw new Error("illegal catch attempt"); }, delegateYield: function delegateYield(iterable, resultName, nextLoc) { return this.delegate = { iterator: values(iterable), resultName: resultName, nextLoc: nextLoc }, "next" === this.method && (this.arg = undefined), ContinueSentinel; } }, exports; }
function asyncGeneratorStep(gen, resolve, reject, _next, _throw, key, arg) { try { var info = gen[key](arg); var value = info.value; } catch (error) { reject(error); return; } if (info.done) { resolve(value); } else { Promise.resolve(value).then(_next, _throw); } }
function _asyncToGenerator(fn) { return function () { var self = this, args = arguments; return new Promise(function (resolve, reject) { var gen = fn.apply(self, args); function _next(value) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "next", value); } function _throw(err) { asyncGeneratorStep(gen, resolve, reject, _next, _throw, "throw", err); } _next(undefined); }); }; }
//import * as msal from '@azure/msal-browser';
//require ("https://alcdn.msauth.net/browser/2.13.1/js/msal-browser.min.js");

// /import { logger } from "./Shared/Helper";

$(function () {
  logger.log('Auth now');

  //logincommon.auth();
  Office.onReady().then(function (info) {
    console.log('office context:', Office.context.ui);
    console.log(info);
    logger.log(info);
    logincommon.auth();
  });
});
var loginfunctions = {};
loginfunctions.loginfn = function () {
  $('#contentdiv').append("<h2>ok</h2>");
  $('#contentdiv').append("<h2>office ready</h2>");
  logincommon.login();
};
loginfunctions.logoutfn = function () {
  Office.onReady().then(function () {
    logincommon.logout();
  });
};
var logincommon = {
  msalConfig: {
    auth: {
      clientId: '121964d0-6350-438e-b38a-54bfedab82bf',
      // Azure application ID         
      authority: 'https://login.microsoftonline.com/common'
    },
    cache: {
      cacheLocation: "localStorage",
      // Needed to avoid "User login is required" error.
      storeAuthStateInCookie: true // Recommended to avoid certain IE/Edge issues.
    }
  },

  _msalInstance: null,
  get msalInstance() {
    if (!logincommon._msalInstance) logincommon._msalInstance = new msal.PublicClientApplication(logincommon.msalConfig);
    return logincommon._msalInstance;
  }
};
logincommon.msalredirecthander = function (res) {
  console.log(res);
  var currentAccounts = logincommon.msalInstance.getAllAccounts();
  if (currentAccounts === null) {
    logger.log("no accounts detected");
  } else if (currentAccounts.length > 1) {
    // Add choose account code here
    logger.log("multiple accounts");
  } else if (currentAccounts.length === 1) {
    username = currentAccounts[0].username;
    logger.log("1 accounts: " + username);
  }
  if (res) {
    authCallback(null, res);
  }
};
logincommon.login = function () {
  try {
    $('#contentdiv').append("<h2>Tryign login</h2>");
    //if(!logincommon.msalInstance) logincommon.msalInstance=new msal.PublicClientApplication(logincommon.msalConfig);
    var loginRequest = {
      scopes: ["User.Read", "Sites.ReadWrite.All"
      //"https://graph.microsoft.com/Sites.ReadWrite.All",
      //"https://incitias.sharepoint.com/AllSites.Manage",
      //"https://incitias.sharepoint.com/User.Read.All"
      ] // optional Array<string>
    };

    logincommon.msalInstance.handleRedirectPromise().then(logincommon.msalredirecthander, function (err) {
      logger.log('logn redirect error: ' + JSON.stringify(err));
    });
    logincommon.msalInstance.loginRedirect(loginRequest);
  } catch (loginerr) {
    $('#contentdiv').append("<h2>Login error: " + JSON.stringify(loginerr) + "</h2>");
    logger.log('Login Error: ' + JSON.stringify(loginerr));
  }
};
logincommon.logout = function () {
  try {
    //if(!logincommon.msalInstance) logincommon.msalInstance=new msal.PublicClientApplication(logincommon.msalConfig);

    var logoutRequest = {
      account: logincommon.msalInstance.getAccountByUsername("admin@sha2013.onmicrosoft.com")
    };
    logincommon.msalInstance.logout(logoutRequest);
  } catch (loginerr) {
    logger.log('Login Error: ' + loginerr);
  }
};
logincommon.auth = /*#__PURE__*/_asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee() {
  var currentAccount, allacocunts, asyncloginRequest, loginRequest, tksil;
  return _regeneratorRuntime().wrap(function _callee$(_context) {
    while (1) switch (_context.prev = _context.next) {
      case 0:
        _context.prev = 0;
        logger.log('Auth now in function');
        allacocunts = logincommon.msalInstance.getAllAccounts(); //getAccountByUsername("admin@sha2013.onmicrosoft.com");
        //var currentAccount = logincommon.msalInstance.getAccountByUsername("ankur.agarwal@incitias.com");
        //logger.log('current account',currentAccount.username);
        if (allacocunts && allacocunts.length > 0) currentAccount = allacocunts[0];else currentAccount = null;
        console.log(currentAccount);
        asyncloginRequest = {
          scopes: ["User.Read", "Sites.ReadWrite.All"
          // "https://incitias.sharepoint.com/AllSites.Manage",
          // "https://incitias.sharepoint.com/User.Read.All"
          ],

          // optional Array<string>
          account: currentAccount,
          //prompt:'consent',
          forceRefresh: false
        };
        loginRequest = {
          scopes: ["User.Read", "Sites.ReadWrite.All"
          // "https://incitias.sharepoint.com/AllSites.Manage",
          //"https://incitias.sharepoint.com/User.Read.All"
          ]
          // prompt:'consent',

          //loginHint: currentAccount.username // optional Array<string>
        };

        logincommon.msalInstance.handleRedirectPromise().then(logincommon.msalredirecthander, function (err) {
          logger.log('logn redirect error: ' + JSON.stringify(err));
        });
        _context.next = 10;
        return logincommon.msalInstance.acquireTokenSilent(asyncloginRequest).then(function (tokenResponse) {
          // Do something with the tokenResponse
          logger.log('got it ' + JSON.stringify(tokenResponse));
          console.log('got it', tokenResponse);
          return tokenResponse;
        }).catch(function (error) {
          logger.log('err: ', error);

          // if (error instanceof InteractionRequiredAuthError)
          //  {
          // fallback to interaction when silent call fails
          return logincommon.msalInstance.acquireTokenRedirect(loginRequest);
          //}
        });
      case 10:
        tksil = _context.sent;
        logger.log(JSON.stringify(tksil));
        console.log('now auth call back', tksil);
        authCallback(null, tksil);
        _context.next = 21;
        break;
      case 16:
        _context.prev = 16;
        _context.t0 = _context["catch"](0);
        console.log('auth error', _context.t0);
        logger.log("Error in Auth");
        logger.log("", _context.t0);
      case 21:
      case "end":
        return _context.stop();
    }
  }, _callee, null, [[0, 16]]);
}));
function authCallback(error, response) {
  console.log('in authcallbacl', response.tokenType);
  if (error) {
    logger.log(error);
    //Office.context.ui.messageParent(JSON.stringify({ status: "failure", result: error }));
  } else {
    if (response.tokenType === "id_token") {
      logger.log("id_token" + response.idToken.rawIdToken);
      //localStorage.setItem("loggedIn", "yes");
    } else {
      logger.log("token type is:" + response.tokenType);
      console.log('token type is', response.tokenType);
      Office.onReady().then(function (info) {
        Office.context.ui.messageParent(JSON.stringify({
          status: "success",
          result: response.accessToken
        }));
      });
    }
  }
}
var logger = {};
logger.log = /*#__PURE__*/function () {
  var _ref2 = _asyncToGenerator( /*#__PURE__*/_regeneratorRuntime().mark(function _callee2(logmsg, logobj) {
    return _regeneratorRuntime().wrap(function _callee2$(_context2) {
      while (1) switch (_context2.prev = _context2.next) {
        case 0:
          try {
            //console.log(logmsg,logobj);    
            if (logobj) logmsg = logmsg + ": " + JSON.stringify(logobj);
            // if (Object.prototype.toString.call(logmsg) === '[object Object]')
            // logmsg=JSON.stringify(logmsg);

            console.log(logmsg);
          } catch (err) {
            console.log('eeror in logger', err);
          }
        case 1:
        case "end":
          return _context2.stop();
      }
    }, _callee2);
  }));
  return function (_x, _x2) {
    return _ref2.apply(this, arguments);
  };
}();
/******/ })()
;
//# sourceMappingURL=login.js.map
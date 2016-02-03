var WacTest;
(function (WacTest) {
    var RichApiHost = (function () {
        function RichApiHost() {
        }
        RichApiHost.prototype.hostName = function () {
            return "";
        };
        RichApiHost.prototype.appId = function () {
            return "";
        };
        RichApiHost.prototype.createInstance = function (typeName) {
            if (typeName == "ExcelApi.TestCaseObject") {
                return new window["ExcelApiWac"].TestCaseObject();
            }
            throw OfficeExtension.WacRuntime.Utility.createInvalidRequestException();
        };
        return RichApiHost;
    })();
    function sendResponseMessage(origin, msgId, result) {
        var responseMsg = { messageId: msgId, messageType: 'RichApiResponse', message: result };
        if (window.parent) {
            window.parent.postMessage(JSON.stringify(responseMsg), origin);
        }
    }
    window.addEventListener("message", function (ev) {
        var msg = JSON.parse(ev.data);
        if (msg && msg.messageId && msg.message && msg.messageType == 'RichApiRequest') {
            var origin = ev.origin;
            var typeRegFunc = window["ExcelApiWac"].TypeRegister.registerTypes;
            var globalObject = new window["ExcelApiWac"].Application();
            var host = new RichApiHost();
            var p = OfficeExtension.WacRuntime.RichApi.executeRichApiRequest(host, typeRegFunc, globalObject, msg.message);
            return p.then(function (result) {
                if (result) {
                    for (var i = 0; i < result.length; i++) {
                        RichApiTest.log.comment(result[i]);
                    }
                }
                sendResponseMessage(origin, msg.messageId, result);
            });
        }
    });
})(WacTest || (WacTest = {}));
var WacTest;
(function (WacTest) {
    function testPromise1() {
        var result = "";
        OfficeExtension.WacRuntime.Utility.createPromiseFromResult(null).then(function () {
            result += "1,";
            return OfficeExtension.WacRuntime.Utility.createPromise(function (resolve, reject) {
                result += "2,";
                resolve(null);
            }).then(function () {
                result += "3,";
            }).then(function () {
                result += "4,";
                return OfficeExtension.WacRuntime.Utility.createPromiseFromResult(null).then(function () {
                    result += "5,";
                });
            });
        }).then(function () {
            result += "6,";
        }).then(function () {
            RichApiTest.log.comment("Result=" + result);
            RichApiTest.log.done(result == "1,2,3,4,5,6,");
        });
    }
    WacTest.testPromise1 = testPromise1;
    function testJsonWriter1() {
        var writer = new OfficeExtension.WacRuntime.JsonWriter();
        writer.startArrayScope();
        writer.startObjectScope();
        writer.writeName("a");
        writer.writeValue("a");
        writer.writeName("b");
        writer.writeValueNull();
        writer.writeName("c");
        writer.startObjectScope();
        writer.endScope();
        writer.writeName("d");
        writer.startArrayScope();
        writer.writeValue(1);
        writer.writeValue(2);
        writer.startObjectScope();
        writer.endScope();
        writer.endScope();
        writer.endScope();
        writer.endScope();
        RichApiTest.log.comment(writer.getJsonString());
    }
    WacTest.testJsonWriter1 = testJsonWriter1;
})(WacTest || (WacTest = {}));

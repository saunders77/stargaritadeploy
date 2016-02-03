var FakeExcelTest;
(function (FakeExcelTest) {
    function testRequestMessage() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        var result1 = range.replaceValue("Hello");
        var result2 = range.replaceValue("HelloWorld");
        ctx.load(range, "Value, RowIndex");
        range.activate();
        var msg = ctx._pendingRequest.buildRequestMessageBody();
        var str = JSON.stringify(msg);
        RichApiTest.log.comment(str);
        RichApiTest.log.done(true);
    }
    FakeExcelTest.testRequestMessage = testRequestMessage;
    function testFakeResponse() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        var result1 = range.replaceValue("Hello");
        var result2 = range.replaceValue("HelloWorld");
        ctx.load(range, "Value, RowIndex");
        range.activate();
        ctx._requestExecutor = new OfficeExtension.FakeResponseRequestExecutor("{\"Results\":[{\"ActionId\":2,\"Value\":true},{\"ActionId\":4,\"Value\":true},{\"ActionId\":6,\"Value\":true},{\"ActionId\":8,\"Value\":true},{\"ActionId\":9,\"Value\":\"Initial Value\"},{\"ActionId\":10,\"Value\":\"Hello\"},{\"ActionId\":11,\"Value\":{\"Value\":\"HelloWorld\",\"RowIndex\":0}},{\"ActionId\":12,\"Value\":null}]}");
        ctx.sync().then(function () {
            RichApiTest.log.comment(result1.value);
            RichApiTest.log.comment(result2.value);
            RichApiTest.log.comment(range.value);
            RichApiTest.log.comment("" + range.rowIndex);
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testFakeResponse = testFakeResponse;
    function testSimpleRequest() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        var result1 = range.replaceValue("Hello");
        var result2 = range.replaceValue("HelloWorld");
        ctx.load(range, "Value, RowIndex");
        range.activate();
        ctx.sync().then(function () {
            RichApiTest.log.comment(result1.value);
            RichApiTest.log.comment(result2.value);
            RichApiTest.log.comment(range.value);
            RichApiTest.log.comment("" + range.rowIndex);
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testSimpleRequest = testSimpleRequest;
    function testUnicode() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var str1 = "\u4e2d\u0386";
        RichApiTest.log.comment(str1);
        RichApiTest.log.comment(JSON.stringify([str1]));
        var str2 = "";
        for (var i = 0; i < 120; i++) {
            str2 = str2 + String.fromCharCode(i + 1);
        }
        RichApiTest.log.comment("Character 1 to 121");
        RichApiTest.log.comment(JSON.stringify([str2]));
        var str3 = "";
        for (var i = 0; i < 1024; i++) {
            str3 = str3 + String.fromCharCode(i + 1);
        }
        RichApiTest.log.comment("Character 1 to 1025");
        RichApiTest.log.comment(JSON.stringify([str3]));
        var range1 = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        range1.value = str1;
        var range2 = ctx.application.activeWorkbook.activeWorksheet.range("A2");
        range2.value = str2;
        var range3 = ctx.application.activeWorkbook.activeWorksheet.range("A3");
        range3.value = str3;
        ctx.load(range1, "Value, RowIndex");
        ctx.load(range2, "Value, RowIndex");
        ctx.load(range3, "Value, RowIndex");
        var msg = ctx._pendingRequest.buildRequestMessageBody();
        var str = JSON.stringify(msg);
        RichApiTest.log.comment("Request Message:");
        RichApiTest.log.comment(str);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Results:");
            RichApiTest.log.comment("Range1.value=" + range1.value);
            RichApiTest.log.comment("Range1.rowIndex=" + range1.rowIndex);
            RichApiTest.log.comment("Range2.value=" + range2.value);
            RichApiTest.log.comment("Range2.rowIndex=" + range2.rowIndex);
            RichApiTest.log.comment("Range3.value=" + range3.value);
            RichApiTest.log.comment("Range3.rowIndex=" + range3.rowIndex);
            RichApiTest.log.done(range1.value == str1 && range2.value == str2 && range3.value == str3);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testUnicode = testUnicode;
    function testWorkbook() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var workbook = ctx.application.activeWorkbook;
        ctx.load(workbook);
        var range = workbook.activeWorksheet.range("A1");
        var result1 = range.replaceValue("Hello");
        var result2 = range.replaceValue("HelloWorld");
        ctx.load(range, "Value");
        range.activate();
        ctx.sync().then(function () {
            RichApiTest.log.comment(result1.value);
            RichApiTest.log.comment(result2.value);
            RichApiTest.log.comment(range.value);
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testWorkbook = testWorkbook;
    function testUpdateValue() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var workbook = ctx.application.activeWorkbook;
        var range = workbook.activeWorksheet.range("A1");
        range.value = "123";
        ctx.load(range, "Value");
        ctx.sync().then(function () {
            RichApiTest.log.comment(range.value);
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testUpdateValue = testUpdateValue;
    function testUpdateText() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var workbook = ctx.application.activeWorkbook;
        var range = workbook.activeWorksheet.range("A1");
        range.text = "abc";
        ctx.load(range, "Text");
        ctx.sync().then(function () {
            RichApiTest.log.comment(range.text);
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testUpdateText = testUpdateText;
    function testObjectNewAndObjectAsParameter() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var testCase = FakeExcelApi.TestCaseObject.newObject(ctx);
        var workbook = ctx.application.activeWorkbook;
        var range = workbook.activeWorksheet.range("A1");
        var result = testCase.calculateAddressAndSaveToRange("One Microsoft Way", "Redmond", range);
        ctx.sync().then(function () {
            RichApiTest.log.comment(result.value);
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testObjectNewAndObjectAsParameter = testObjectNewAndObjectAsParameter;
    function testArrayValue() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var sheet = ctx.application.activeWorkbook.activeWorksheet;
        var range = sheet.range("A1");
        range.value = ['Hello', 123, true];
        ctx.load(range);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Succeeded");
            RichApiTest.log.comment("Range.Value=" + JSON.stringify(range.value));
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testArrayValue = testArrayValue;
    function testValue2DArray() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var sheet = ctx.application.activeWorkbook.activeWorksheet;
        var range = sheet.range("A1");
        range.valueArray = [['Hello', 123, true], ['World', 456, false]];
        ctx.load(range);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Succeeded");
            RichApiTest.log.comment("Range.ValueArray=" + JSON.stringify(range.valueArray));
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testValue2DArray = testValue2DArray;
    function testText2DArray() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var sheet = ctx.application.activeWorkbook.activeWorksheet;
        var range = sheet.range("A1");
        range.textArray = [["Abc", "Def"], ["Seattle", "Redmond"]];
        ctx.load(range);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Succeeded");
            RichApiTest.log.comment("Range.TextArray=" + JSON.stringify(range.textArray));
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testText2DArray = testText2DArray;
    function testValueArray2() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var sheet = ctx.application.activeWorkbook.activeWorksheet;
        var range = sheet.range("A1");
        range.valueArray2 = ['A', 'B', 'C'];
        ctx.load(range);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Succeeded");
            RichApiTest.log.comment("Range.ValueArray2=" + JSON.stringify(range.valueArray2));
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testValueArray2 = testValueArray2;
    function testErrorWorksheet() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var sheet = ctx.application.testWorkbook.errorWorksheet;
        ctx.load(sheet);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Succeeded, but expect failure");
            RichApiTest.log.done(false);
        }, function (errorInfo) {
            RichApiTest.log.comment("Expected E_CHANGED_STATE. ErrorCode=" + errorInfo.code);
            RichApiTest.log.comment("ErrorMessage=" + errorInfo.message);
            RichApiTest.log.comment("ErrorLocation=" + errorInfo.debugInfo.errorLocation);
            RichApiTest.log.done(errorInfo.code == FakeExcelApi.ErrorCodes.conflict && errorInfo.debugInfo.errorLocation == "TestWorkbook.errorWorksheet");
        });
    }
    FakeExcelTest.testErrorWorksheet = testErrorWorksheet;
    function testErrorWorksheet2() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var sheet = ctx.application.testWorkbook.errorWorksheet2;
        ctx.load(sheet);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Succeeded, but expect failure");
            RichApiTest.log.done(false);
        }, function (errorInfo) {
            RichApiTest.log.comment("Expected E_CHANGED_STATE. ErrorCode=" + errorInfo.code);
            RichApiTest.log.comment("ErrorMessage=" + errorInfo.message);
            RichApiTest.log.comment("ErrorLocation=" + errorInfo.debugInfo.errorLocation);
            RichApiTest.log.done(errorInfo.code == FakeExcelApi.ErrorCodes.conflict2 && errorInfo.debugInfo.errorLocation == "TestWorkbook.errorWorksheet2");
        });
    }
    FakeExcelTest.testErrorWorksheet2 = testErrorWorksheet2;
    function testErrorWorksheet2_rest() {
        jQuery.ajax(OfficeExtension.Constants.localDocumentApiPrefix + "testWorkbook/errorWorksheet2").done(function (data) {
            RichApiTest.log.comment(data);
            RichApiTest.log.comment("Expect failure");
            RichApiTest.log.done(false);
        }).fail(function (jqXHR) {
            RichApiTest.log.comment("StatusCode=" + jqXHR.status);
            RichApiTest.log.comment(jqXHR.responseText);
            RichApiTest.log.done(jqXHR.status == 409);
        });
        ;
    }
    FakeExcelTest.testErrorWorksheet2_rest = testErrorWorksheet2_rest;
    function testErrorMethodAccessDenied() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var result = ctx.application.testWorkbook.errorMethod(FakeExcelApi.ErrorMethodType.accessDenied);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Succeeded, but expect failure");
            RichApiTest.log.done(false);
        }, function (errorInfo) {
            RichApiTest.log.comment("Expected E_ACCESSDENIED. ErrorCode=" + errorInfo.code);
            RichApiTest.log.comment("ErrorMessage=" + errorInfo.message);
            RichApiTest.log.comment("ErrorLocation=" + errorInfo.debugInfo.errorLocation);
            RichApiTest.log.done(errorInfo.code == FakeExcelApi.ErrorCodes.accessDenied && errorInfo.debugInfo.errorLocation == "TestWorkbook.errorMethod");
        });
    }
    FakeExcelTest.testErrorMethodAccessDenied = testErrorMethodAccessDenied;
    function testErrorMethodBounds() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var result = ctx.application.testWorkbook.errorMethod(FakeExcelApi.ErrorMethodType.bounds);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Succeeded, but expect failure");
            RichApiTest.log.done(false);
        }, function (errorInfo) {
            RichApiTest.log.comment("Expected E_Bounds. ErrorCode=" + errorInfo.code);
            RichApiTest.log.comment("ErrorMessage=" + errorInfo.message);
            RichApiTest.log.comment("ErrorLocation=" + errorInfo.debugInfo.errorLocation);
            RichApiTest.log.done(errorInfo.code == FakeExcelApi.ErrorCodes.outOfRange && errorInfo.debugInfo.errorLocation == "TestWorkbook.errorMethod");
        });
    }
    FakeExcelTest.testErrorMethodBounds = testErrorMethodBounds;
    function testErrorMethod2Abort() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var result = ctx.application.testWorkbook.errorMethod2(FakeExcelApi.ErrorMethodType.abort);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Succeeded, but expect failure");
            RichApiTest.log.done(false);
        }, function (errorInfo) {
            RichApiTest.log.comment("Expected E_Bounds. ErrorCode=" + errorInfo.code);
            RichApiTest.log.comment("ErrorMessage=" + errorInfo.message);
            RichApiTest.log.done(errorInfo.code == FakeExcelApi.ErrorCodes.aborted2);
        });
    }
    FakeExcelTest.testErrorMethod2Abort = testErrorMethod2Abort;
    function testErrorMethod2AccessDenied() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var result = ctx.application.testWorkbook.errorMethod2(FakeExcelApi.ErrorMethodType.accessDenied);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Succeeded, but expect failure");
            RichApiTest.log.done(false);
        }, function (errorInfo) {
            RichApiTest.log.comment("Expected E_Bounds. ErrorCode=" + errorInfo.code);
            RichApiTest.log.comment("ErrorMessage=" + errorInfo.message);
            RichApiTest.log.done(errorInfo.code == FakeExcelApi.ErrorCodes.accessDenied2);
        });
    }
    FakeExcelTest.testErrorMethod2AccessDenied = testErrorMethod2AccessDenied;
    function testObjectCount() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        ctx.load(range);
        var countResult = ctx.application.testWorkbook.getObjectCount();
        ctx.sync().then(function () {
            RichApiTest.log.comment("Range: row=" + range.rowIndex);
            RichApiTest.log.comment("ObjectCount=" + countResult.value);
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testObjectCount = testObjectCount;
    function testObjectCountSimple() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var book = ctx.application.activeWorkbook;
        ctx.load(book);
        var countResult = ctx.application.testWorkbook.getObjectCount();
        ctx.sync().then(function () {
            RichApiTest.log.comment("ObjectCount=" + countResult.value);
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testObjectCountSimple = testObjectCountSimple;
    function testTraceFailure() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var workbook = ctx.application.activeWorkbook;
        ctx.trace("LoadingWorkbook");
        ctx.load(workbook);
        ctx.trace("LoadedWorkbook");
        var result = ctx.application.testWorkbook.errorMethod(FakeExcelApi.ErrorMethodType.accessDenied);
        var range = workbook.activeWorksheet.range("A1");
        var resultRangeNewValue = range.replaceValue("Hello");
        ctx.trace("ReplacedValue with Hello");
        ctx.sync().then(function () {
            RichApiTest.log.comment("Succeeded, but expect failure");
            RichApiTest.log.done(false);
        }, function (errorInfo) {
            RichApiTest.log.comment("Expected E_ACCESSDENIED. ErrorCode=" + errorInfo.code);
            RichApiTest.log.comment("TraceInfos:");
            RichApiTest.log.comment(JSON.stringify(errorInfo.traceMessages));
            RichApiTest.log.done(true);
        });
    }
    FakeExcelTest.testTraceFailure = testTraceFailure;
    function verifySheet(sheet, index) {
        var id = 2000 + index;
        var success = true;
        if (sheet._Id != id) {
            RichApiTest.log.comment("Sheet.Id is not " + id);
            success = false;
        }
        var name = "Sheet" + id;
        if (sheet.name != name) {
            RichApiTest.log.comment("Sheet.name is not " + name);
            success = false;
        }
        if (sheet.calculatedName != name) {
            RichApiTest.log.comment("Sheet.calculatedName is not " + name);
            success = false;
        }
        return success;
    }
    function testWorksheetCollection() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var workbook = ctx.application.activeWorkbook;
        var oneSheet = workbook.sheets.getItem("Sheet2004");
        ctx.load(workbook.sheets);
        ctx.load(oneSheet);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Count=" + workbook.sheets.items.length);
            var success = true;
            if (workbook.sheets.items.length < 5) {
                RichApiTest.log.comment("items.length < 5");
                success = false;
            }
            for (var i = 0; i < workbook.sheets.items.length; i++) {
                RichApiTest.log.comment("Sheet" + i + ": Name=" + workbook.sheets.items[i].name);
                RichApiTest.log.comment("Sheet" + i + ": Id=" + workbook.sheets.items[i]._Id);
                RichApiTest.log.comment("Sheet" + i + ": CalculatedName=" + workbook.sheets.items[i].calculatedName);
                success = success && verifySheet(workbook.sheets.items[i], i);
            }
            RichApiTest.log.comment("OneSheet: Name=" + oneSheet.name);
            RichApiTest.log.done(success);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testWorksheetCollection = testWorksheetCollection;
    function testWorksheetCollectionThenOneItem() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var workbook = ctx.application.activeWorkbook;
        ctx.load(workbook.sheets);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Count=" + workbook.sheets.items.length);
            var success = true;
            if (workbook.sheets.items.length < 5) {
                RichApiTest.log.comment("items.length < 5");
                success = false;
            }
            var name1 = workbook.sheets.items[3].name;
            var oneSheet = workbook.sheets.items[3];
            ctx.load(oneSheet);
            ctx.sync().then(function () {
                RichApiTest.log.comment("oneSheet.name = " + oneSheet.name);
                if (oneSheet.name != name1) {
                    RichApiTest.log.comment("The oneSheet.name != name1");
                    success = false;
                }
                RichApiTest.log.done(success);
            }, FakeExcelTest.reportError);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testWorksheetCollectionThenOneItem = testWorksheetCollectionThenOneItem;
    function testWorksheetCollectionActiveWorksheetActiveCell() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var workbook = ctx.application.activeWorkbook;
        var sheet = workbook.sheets.getActiveWorksheetInvalidAfterRequest();
        // The sheet's object path will be invlaidated, then fixed up
        var range = sheet.getActiveCell();
        ctx.sync().then(function () {
            // This is to test whether range still have valid object path as the sheet's object path was fixed.
            ctx.load(range);
            ctx.sync().then(function () {
                RichApiTest.log.comment("range.value = " + range.value);
                RichApiTest.log.done(true);
            }, FakeExcelTest.reportError);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testWorksheetCollectionActiveWorksheetActiveCell = testWorksheetCollectionActiveWorksheetActiveCell;
    function testWorksheetCollectionAdd() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var workbook = ctx.application.activeWorkbook;
        var sheet = workbook.sheets.add("SomeName");
        ctx.sync().then(function () {
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testWorksheetCollectionAdd = testWorksheetCollectionAdd;
    function testWorksheetCollection_rest() {
        jQuery.ajax(OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/sheets").done(function (data) {
            RichApiTest.log.comment(data);
            var v = JSON.parse(data);
            var success = true;
            RichApiTest.log.comment("sheets.length = " + v.value.length);
            if (v.value.length < 5) {
                RichApiTest.log.comment("v.value.length < 5");
                success = false;
            }
            RichApiTest.log.done(success);
        }).fail(function (jqXHR) {
            RichApiTest.log.comment("StatusCode=" + jqXHR.status);
            RichApiTest.log.done(false);
        });
        ;
    }
    FakeExcelTest.testWorksheetCollection_rest = testWorksheetCollection_rest;
    function worksheetCollectionAddMethod_rest(appendAdd, appendDollar) {
        var bodyObj = {};
        bodyObj.name = "SomeName";
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/sheets",
            method: RichApiTest.RestUtility.httpMethodPost,
            body: JSON.stringify(bodyObj)
        };
        if (appendAdd) {
            if (appendDollar) {
                request.url = request.url + "/$/add";
            }
            else {
                request.url = request.url + "/add";
            }
        }
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusCreated);
            var sheet = JSON.parse(resp.body);
            if (!sheet) {
                throw "Cannot parse sheet";
            }
            RichApiTest.log.comment("odata.type=" + sheet["@odata.type"]);
            if (sheet["@odata.type"] != "ExcelApi.Worksheet") {
                throw "odata.type is not ExcelApi.Worksheet";
            }
            if (sheet["_Id"]) {
                throw "Should not have _Id property";
            }
            var id = sheet["@odata.id"];
            RichApiTest.log.comment("odata.id=" + id);
            var expectedIdPrefix = "activeWorkbook/sheets(";
            if (id.substr(0, expectedIdPrefix.length) != expectedIdPrefix) {
                throw "@odata.id is not " + expectedIdPrefix;
            }
            return id;
        }).then(function (odataId) {
            var request = {
                url: OfficeExtension.Constants.localDocumentApiPrefix + odataId,
                method: RichApiTest.RestUtility.httpMethodGet
            };
            return RichApiTest.RestUtility.invoke(request);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            RichApiTest.log.done(true);
        }).catch(function (err) {
            RichApiTest.log.fail(JSON.stringify(err));
        });
    }
    FakeExcelTest.worksheetCollectionAddMethod_rest = worksheetCollectionAddMethod_rest;
    function testWorksheetCollectionAdd_rest() {
        FakeExcelTest.worksheetCollectionAddMethod_rest(false, false);
    }
    FakeExcelTest.testWorksheetCollectionAdd_rest = testWorksheetCollectionAdd_rest;
    function testWorksheetCollectionAddMethod_rest() {
        FakeExcelTest.worksheetCollectionAddMethod_rest(true, false);
    }
    FakeExcelTest.testWorksheetCollectionAddMethod_rest = testWorksheetCollectionAddMethod_rest;
    function testWorksheetCollectionAddMethodDollar_rest() {
        FakeExcelTest.worksheetCollectionAddMethod_rest(true, true);
    }
    FakeExcelTest.testWorksheetCollectionAddMethodDollar_rest = testWorksheetCollectionAddMethodDollar_rest;
    function testActiveCell() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var workbook = ctx.application.activeWorkbook;
        var oneSheet = workbook.sheets.getItem("Sheet2004");
        var cell1 = oneSheet.getActiveCell();
        var cell2 = oneSheet.getActiveCell();
        ctx.load(cell1);
        ctx.load(cell2);
        ctx.sync().then(function () {
            RichApiTest.log.comment("cell1: value=" + cell1.value);
            RichApiTest.log.comment("cell2: value=" + cell2.value);
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testActiveCell = testActiveCell;
    function testActiveCellInvalidAfterRequest() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var workbook = ctx.application.activeWorkbook;
        var oneSheet = workbook.sheets.getItem("Sheet2004");
        var cell1 = oneSheet.getActiveCellInvalidAfterRequest();
        var cell2 = oneSheet.getActiveCellInvalidAfterRequest();
        ctx.load(cell1);
        ctx.load(cell2);
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("cell1: value=" + cell1.value);
            RichApiTest.log.comment("cell2: value=" + cell2.value);
            try {
                ctx.load(cell1);
                RichApiTest.log.comment("Does not get expected exception in ctx.load(cell1)");
                success = false;
            }
            catch (ex) {
                if (ex instanceof OfficeExtension.Error && ex.code == "InvalidObjectPath") {
                    RichApiTest.log.comment("Expect exception in ctx.load(): " + ex);
                }
                else {
                    success = false;
                }
            }
            try {
                oneSheet.someRangeOperation("Test", cell2);
                RichApiTest.log.comment("Does not get expected exception in oneSheet.someRangeOperation(cell2)");
                success = false;
            }
            catch (ex) {
                if (ex instanceof OfficeExtension.Error && ex.code == "InvalidObjectPath") {
                    RichApiTest.log.comment("Expect exception in oneSheet.someRangeOperation: " + ex);
                }
                else {
                    success = false;
                }
            }
            RichApiTest.log.done(success);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testActiveCellInvalidAfterRequest = testActiveCellInvalidAfterRequest;
    function logError(error) {
        RichApiTest.log.comment(JSON.stringify(error));
        if (error instanceof OfficeExtension.Error) {
            RichApiTest.log.comment("FakeExcelApi Host Error = " + error.code);
            RichApiTest.log.comment("Location=" + error.debugInfo.errorLocation);
        }
        else {
            RichApiTest.log.comment("Error = " + error.message);
        }
    }
    FakeExcelTest.logError = logError;
    function reportError(error) {
        logError(error);
        RichApiTest.log.done(false);
    }
    FakeExcelTest.reportError = reportError;
    function reportAjaxError(jqXHR) {
        RichApiTest.log.comment("StatusCode=" + jqXHR.status);
        RichApiTest.log.done(false);
    }
    FakeExcelTest.reportAjaxError = reportAjaxError;
    function testGetSheets_rest() {
        jQuery.ajax(OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/sheets").done(function (data) {
            RichApiTest.log.comment(data);
            RichApiTest.log.done(true);
        }).fail(function (jqXHR) {
            RichApiTest.log.comment("StatusCode=" + jqXHR.status);
            RichApiTest.log.done(false);
        });
        ;
    }
    FakeExcelTest.testGetSheets_rest = testGetSheets_rest;
    function testPatchRange_rest() {
        var option = {};
        option.url = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/sheets(2000)/Range('A1')";
        option.type = "PATCH";
        option.data = "{\"text\": \"NewText\", \"value\": \"New Value\"}";
        jQuery.ajax(option).done(function (data) {
            RichApiTest.log.comment(data);
            RichApiTest.log.done(true);
        }).fail(function (jqXHR) {
            RichApiTest.log.comment("StatusCode=" + jqXHR.status);
            RichApiTest.log.done(false);
        });
    }
    FakeExcelTest.testPatchRange_rest = testPatchRange_rest;
    function testPatchRange2_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/sheets(2000)/rAnge('A''1')",
            method: RichApiTest.RestUtility.httpMethodPatch,
            body: "{\"text\": \"NewText\", \"value\": \"New Value\"}",
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            var range = JSON.parse(resp.body);
            var value = range["value"];
            if (value != "New Value") {
                throw "value is not New Value";
            }
            if (range["@odata.type"] != "ExcelApi.Range") {
                throw "odata.type is not ExcelApi.Range";
            }
            var expectedId = "activeWorkbook/sheets(2000)/range(address='A''1')";
            var id = range["@odata.id"];
            RichApiTest.log.comment("odata.id=" + id);
            id = decodeURI(id);
            RichApiTest.log.comment("decodeURI(odata.id)=" + id);
            if (id != expectedId) {
                throw "odata.id is not " + expectedId;
            }
            RichApiTest.log.done(true);
        }).catch(function (err) {
            RichApiTest.log.fail(JSON.stringify(err));
        });
    }
    FakeExcelTest.testPatchRange2_rest = testPatchRange2_rest;
    function testGetActiveSheet() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var workbook = ctx.application.testWorkbook;
        var sheet = workbook.getActiveWorksheet();
        ctx.load(sheet);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Name=" + sheet.name);
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testGetActiveSheet = testGetActiveSheet;
    function testGetActiveSheet_rest() {
        jQuery.ajax(OfficeExtension.Constants.localDocumentApiPrefix + "testWorkbook/activeworksheet").done(function (data) {
            RichApiTest.log.comment(data);
            RichApiTest.log.done(true);
        }).fail(function (jqXHR) {
            RichApiTest.log.comment("StatusCode=" + jqXHR.status);
            RichApiTest.log.done(false);
        });
    }
    FakeExcelTest.testGetActiveSheet_rest = testGetActiveSheet_rest;
    function testGetChart() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var chart = ctx.application.activeWorkbook.charts.getItem("Chart1001");
        ctx.load(chart);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Name=" + chart.name);
            RichApiTest.log.comment("ChartType=" + chart.chartType);
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testGetChart = testGetChart;
    function testGetCharts() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var charts = ctx.application.activeWorkbook.charts;
        ctx.load(charts);
        ctx.sync().then(function () {
            for (var i = 0; i < charts.items.length; i++) {
                RichApiTest.log.comment("Name=" + charts.items[i].name);
                RichApiTest.log.comment("ChartType=" + charts.items[i].chartType);
            }
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testGetCharts = testGetCharts;
    function testGetChartsAndThenUpdate() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var charts = ctx.application.activeWorkbook.charts;
        ctx.load(charts);
        ctx.sync().then(function () {
            for (var i = 0; i < charts.items.length; i++) {
                RichApiTest.log.comment("Name=" + charts.items[i].name);
                RichApiTest.log.comment("ChartType=" + charts.items[i].chartType);
                charts.items[i].title = charts.items[i].title + "-New";
            }
            ctx.sync().then(function () {
                RichApiTest.log.done(true);
            }, FakeExcelTest.reportError);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testGetChartsAndThenUpdate = testGetChartsAndThenUpdate;
    function testCreateChart() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var chart = ctx.application.activeWorkbook.charts.add("Chart4", "bar");
        ctx.load(chart);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Name=" + chart.name);
            RichApiTest.log.comment("ChartType=" + chart.chartType);
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testCreateChart = testCreateChart;
    function testCreateChartAndThenUpdateChart() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var chart = ctx.application.activeWorkbook.charts.add("Chart5", FakeExcelApi.ChartType._3DBar);
        ctx.load(chart);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Name=" + chart.name);
            RichApiTest.log.comment("ChartType=" + chart.chartType);
            chart.title = "HelloWorld";
            ctx.load(chart);
            ctx.sync().then(function () {
                RichApiTest.log.comment("Title=" + chart.title);
                RichApiTest.log.done(true);
            }, FakeExcelTest.reportError);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testCreateChartAndThenUpdateChart = testCreateChartAndThenUpdateChart;
    function testCreateChartAndThenUpdateChart2() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var chart = ctx.application.activeWorkbook.charts.add("Chart6", FakeExcelApi.ChartType._3DBar);
        // Intentionally do not call ctx.load(chart);
        ctx.sync().then(function () {
            chart.title = "HelloWorld";
            ctx.load(chart);
            ctx.sync().then(function () {
                RichApiTest.log.comment("Title=" + chart.title);
                RichApiTest.log.done(true);
            }, FakeExcelTest.reportError);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testCreateChartAndThenUpdateChart2 = testCreateChartAndThenUpdateChart2;
    function testUpdateChart() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var chart = ctx.application.activeWorkbook.charts.getItem("Chart1000");
        chart.title = "New Title";
        chart.chartType = FakeExcelApi.ChartType.line;
        ctx.load(chart);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Name=" + chart.name);
            RichApiTest.log.comment("ChartType=" + chart.chartType);
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testUpdateChart = testUpdateChart;
    function testDeleteChart() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var chart = ctx.application.activeWorkbook.charts.getItem("Chart1000");
        chart.delete();
        ctx.sync().then(function () {
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testDeleteChart = testDeleteChart;
    function testGetCharts_rest() {
        jQuery.ajax(OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/charts").done(function (data) {
            RichApiTest.log.comment(data);
            RichApiTest.log.done(true);
        }).fail(function (jqXHR) {
            RichApiTest.log.comment("StatusCode=" + jqXHR.status);
            RichApiTest.log.done(false);
        });
    }
    FakeExcelTest.testGetCharts_rest = testGetCharts_rest;
    function testGetChartByName_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/charts('Chart1002')",
            method: RichApiTest.RestUtility.httpMethodGet,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            var result = JSON.parse(resp.body);
            var id = result["@odata.id"];
            var expectedId = "activeWorkbook/charts(1002)";
            if (id != expectedId) {
                throw "Not get expected id " + expectedId;
            }
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testGetChartByName_rest = testGetChartByName_rest;
    function testGetChartByName2_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/charts('Chart1002')/ChartType",
            method: RichApiTest.RestUtility.httpMethodGet,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            var result = JSON.parse(resp.body);
            var chartType = result["value"];
            if (chartType !== FakeExcelApi.ChartType.bar) {
                throw "Not get expected chartType " + chartType;
            }
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testGetChartByName2_rest = testGetChartByName2_rest;
    function testGetChartByInt_rest() {
        jQuery.ajax(OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/charts(1001)").done(function (data) {
            RichApiTest.log.comment(data);
            RichApiTest.log.done(true);
        }).fail(function (jqXHR) {
            RichApiTest.log.comment("StatusCode=" + jqXHR.status);
            RichApiTest.log.done(false);
        });
    }
    FakeExcelTest.testGetChartByInt_rest = testGetChartByInt_rest;
    function testPatchChart_rest() {
        var option = {};
        option.url = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/charts(1000)";
        option.type = "PATCH";
        option.data = "{\"title\": \"NewTitle\", \"chartType\": \"Line\"}";
        jQuery.ajax(option).done(function (data) {
            RichApiTest.log.comment(data);
            RichApiTest.log.done(true);
        }).fail(function (jqXHR) {
            RichApiTest.log.comment("StatusCode=" + jqXHR.status);
            RichApiTest.log.done(false);
        });
    }
    FakeExcelTest.testPatchChart_rest = testPatchChart_rest;
    function testGetChartByType_rest() {
        jQuery.ajax(OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/chartByType('Bar')").done(function (data) {
            RichApiTest.log.comment(data);
            RichApiTest.log.done(true);
        }).fail(function (jqXHR) {
            RichApiTest.log.comment("StatusCode=" + jqXHR.status);
            RichApiTest.log.done(false);
        });
    }
    FakeExcelTest.testGetChartByType_rest = testGetChartByType_rest;
    function testGetChartByType2_rest() {
        jQuery.ajax(OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/chartByType(chartType='Bar')").done(function (data) {
            RichApiTest.log.comment(data);
            RichApiTest.log.done(true);
        }).fail(function (jqXHR) {
            RichApiTest.log.comment("StatusCode=" + jqXHR.status);
            RichApiTest.log.done(false);
        });
    }
    FakeExcelTest.testGetChartByType2_rest = testGetChartByType2_rest;
    function testGetChartByTypeTitle_rest() {
        jQuery.ajax(OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/chartByTypeTitle(chartType='Bar', title='Abc')").done(function (data) {
            RichApiTest.log.comment(data);
            RichApiTest.log.done(true);
        }).fail(function (jqXHR) {
            RichApiTest.log.comment("StatusCode=" + jqXHR.status);
            RichApiTest.log.done(false);
        });
    }
    FakeExcelTest.testGetChartByTypeTitle_rest = testGetChartByTypeTitle_rest;
    function testCreateChart_rest() {
        var option = {};
        option.url = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/charts";
        option.type = "POST";
        option.data = "{\"title\": \"NewTitle\", \"chartType\": \"Bar\", \"name\": \"SomeName\"}";
        jQuery.ajax(option).done(function (data) {
            var success = true;
            RichApiTest.log.comment(data);
            var jsonObj = JSON.parse(data);
            if (jsonObj) {
                if (typeof (jsonObj.Title) != "undefined") {
                    success = false;
                }
                if (jsonObj.title != "NewTitle") {
                    success = false;
                }
                if (jsonObj.chartType != "Bar") {
                    success = false;
                }
            }
            else {
                success = false;
            }
            RichApiTest.log.done(success);
        }).fail(function (jqXHR) {
            RichApiTest.log.comment("StatusCode=" + jqXHR.status);
            RichApiTest.log.done(false);
        });
    }
    FakeExcelTest.testCreateChart_rest = testCreateChart_rest;
    function testDeleteChart_rest() {
        var option = {};
        option.url = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/charts(1004)";
        option.type = "DELETE";
        jQuery.ajax(option).done(function (data, textStatus, jqXHR) {
            RichApiTest.log.comment("StatusCode=" + jqXHR.status);
            RichApiTest.log.done(true);
        }).fail(function (jqXHR) {
            RichApiTest.log.comment("StatusCode=" + jqXHR.status);
            RichApiTest.log.done(false);
        });
    }
    FakeExcelTest.testDeleteChart_rest = testDeleteChart_rest;
    function testGetAddChart_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/charts/add",
            method: RichApiTest.RestUtility.httpMethodGet,
            body: "{\"title\": \"NewTitle\", \"chartType\": \"Bar\", \"name\": \"SomeName\"}"
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusBadRequest);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testGetAddChart_rest = testGetAddChart_rest;
    function testPatchAddChart_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/charts/add",
            method: RichApiTest.RestUtility.httpMethodPatch,
            body: "{\"title\": \"NewTitle\", \"chartType\": \"Bar\", \"name\": \"SomeName\"}"
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusBadRequest);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testPatchAddChart_rest = testPatchAddChart_rest;
    function testDeleteAddChart_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/charts/add",
            method: RichApiTest.RestUtility.httpMethodDelete,
            body: "{\"title\": \"NewTitle\", \"chartType\": \"Bar\", \"name\": \"SomeName\"}"
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusBadRequest);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testDeleteAddChart_rest = testDeleteAddChart_rest;
    function testSomeAction_rest() {
        var option = {};
        option.url = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/someaction";
        option.type = "POST";
        option.data = "{\"intVal\": 345, \"strVal\": \"Redmond\", \"enumVal\": \"Line\"}";
        jQuery.ajax(option).done(function (data) {
            RichApiTest.log.comment(data);
            RichApiTest.log.done(true);
        }).fail(function (jqXHR) {
            RichApiTest.log.comment("StatusCode=" + jqXHR.status);
            RichApiTest.log.done(false);
        });
    }
    FakeExcelTest.testSomeAction_rest = testSomeAction_rest;
    function testSomeAction2_rest() {
        var option = {};
        option.url = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/someaction(intVal=345, strVal='Redmond', enumVal='Line')";
        option.type = "POST";
        jQuery.ajax(option).done(function (data) {
            RichApiTest.log.comment(data);
            RichApiTest.log.done(true);
        }).fail(function (jqXHR) {
            RichApiTest.log.comment("StatusCode=" + jqXHR.status);
            RichApiTest.log.done(false);
        });
    }
    FakeExcelTest.testSomeAction2_rest = testSomeAction2_rest;
    function testNullable1() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var chart0 = ctx.application.activeWorkbook.charts.getItem("Chart1000");
        ctx.load(chart0);
        var chart1 = ctx.application.activeWorkbook.charts.getItem("Chart1001");
        ctx.load(chart1);
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("Chart0.nullableChartType=" + chart0.nullableChartType);
            if (chart0.nullableChartType !== null) {
                RichApiTest.log.comment("Chart0.nullableChartType should be null");
                success = false;
            }
            RichApiTest.log.comment("Chart1.nullableChartType=" + chart1.nullableChartType);
            if (chart1.nullableChartType !== FakeExcelApi.ChartType.pie) {
                RichApiTest.log.comment("Chart1.nullableChartType should be pie");
                success = false;
            }
            RichApiTest.log.comment("Chart0.nullableShowLabel=" + chart0.nullableShowLabel);
            if (chart0.nullableShowLabel !== null) {
                RichApiTest.log.comment("Chart0.nullableShowLabel should be null");
                success = false;
            }
            RichApiTest.log.comment("Chart1.nullableShowLabel=" + chart1.nullableShowLabel);
            if (chart1.nullableShowLabel !== true) {
                RichApiTest.log.comment("Chart1.nullableShowLabel should be true");
                success = false;
            }
            RichApiTest.log.done(success);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testNullable1 = testNullable1;
    function testNullable2() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var r1 = ctx.application.testWorkbook.testNullableInputValue(null, null);
        var r2 = ctx.application.testWorkbook.testNullableInputValue(FakeExcelApi.ChartType.pie, true);
        var r3 = ctx.application.testWorkbook.getNullableBoolValue(true);
        var r4 = ctx.application.testWorkbook.getNullableBoolValue(false);
        var r5 = ctx.application.testWorkbook.getNullableEnumValue(true);
        var r6 = ctx.application.testWorkbook.getNullableEnumValue(false);
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("r1.value=" + r1.value);
            RichApiTest.log.comment("r2.value=" + r2.value);
            RichApiTest.log.comment("r3.value=" + r3.value);
            if (r3.value !== null) {
                RichApiTest.log.comment("getNullableBoolValue(true) should return null");
                success = false;
            }
            RichApiTest.log.comment("r4.value=" + r4.value);
            if (r4.value !== true) {
                RichApiTest.log.comment("getNullableBoolValue(false) should return true");
                success = false;
            }
            RichApiTest.log.comment("r5.value=" + r5.value);
            if (r5.value !== null) {
                RichApiTest.log.comment("getNullableEnumValue(true) should return null");
                success = false;
            }
            RichApiTest.log.comment("r6.value=" + r6.value);
            if (r6.value != FakeExcelApi.ChartType.pie) {
                RichApiTest.log.comment("getNullableEnumValue(false) should return pie");
                success = false;
            }
            RichApiTest.log.done(success);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testNullable2 = testNullable2;
    function testKeepReference0() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var r = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        r.text = "New Text";
        ctx.sync().then(function () {
            try {
                ctx.load(r);
                RichApiTest.log.fail("Should get exception as keepReference() is not called.");
                return;
            }
            catch (ex) {
                if (ex instanceof OfficeExtension.Error && ex.code == "InvalidObjectPath") {
                    RichApiTest.log.pass("Get expected exception " + ex);
                    return;
                }
            }
            RichApiTest.log.fail("Should not have reached here");
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testKeepReference0 = testKeepReference0;
    function testKeepReference1() {
        var newText = "Hello, Keep Reference";
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var count0 = ctx.application.testWorkbook.getCachedObjectCount();
        var r = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        r.text = newText;
        ctx.trackedObjects.add(r);
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("StartCachedObjectCount = " + count0.value);
            ctx.load(r);
            var count1 = ctx.application.testWorkbook.getCachedObjectCount();
            ctx.sync().then(function () {
                RichApiTest.log.comment("CachedObjectCount = " + count1.value);
                RichApiTest.log.comment("Range.text=" + r.text);
                if (r.text != newText) {
                    success = false;
                }
                ctx.trackedObjects.remove(r);
                var count2 = ctx.application.testWorkbook.getCachedObjectCount();
                ctx.sync().then(function () {
                    RichApiTest.log.comment("EndCachedObjectCount = " + count2.value);
                    if (count2.value != count0.value) {
                        RichApiTest.log.comment("The reference count does not match");
                        success = false;
                    }
                    RichApiTest.log.done(success);
                }, FakeExcelTest.reportError);
            }, FakeExcelTest.reportError);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testKeepReference1 = testKeepReference1;
    function testKeepReference2() {
        var newText = "Hello, Keep Reference2";
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var count0 = ctx.application.testWorkbook.getCachedObjectCount();
        var r = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        r.text = newText;
        ctx.trackedObjects.add(r);
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("SartCachedObjectCount = " + count0.value);
            RichApiTest.log.comment("Add reference again");
            ctx.trackedObjects.add(r);
            ctx.load(r);
            var count1 = ctx.application.testWorkbook.getCachedObjectCount();
            ctx.sync().then(function () {
                RichApiTest.log.comment("CachedObjectCount = " + count1.value);
                RichApiTest.log.comment("Range.text=" + r.text);
                if (r.text != newText) {
                    success = false;
                }
                ctx.trackedObjects.remove(r);
                var count2 = ctx.application.testWorkbook.getCachedObjectCount();
                ctx.sync().then(function () {
                    RichApiTest.log.comment("EndCachedObjectCount = " + count2.value);
                    if (count2.value != count0.value) {
                        RichApiTest.log.comment("The reference count does not match");
                        success = false;
                    }
                    RichApiTest.log.done(success);
                }, FakeExcelTest.reportError);
            }, FakeExcelTest.reportError);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testKeepReference2 = testKeepReference2;
    function testKeepReferenceNoLoad() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var countStart = ctx.application.testWorkbook.getCachedObjectCount();
        var r = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        ctx.trackedObjects.add(r);
        // intentionally do not call ctx.load(r);
        ctx.sync().then(function () {
            RichApiTest.log.comment("CountStart=" + countStart.value);
            RichApiTest.log.comment("ReferenceId=" + r._ReferenceId);
            var success = true;
            if (r._ReferenceId == null || r._ReferenceId.length == 0) {
                success = false;
            }
            ctx.trackedObjects.remove(r);
            var countEnd = ctx.application.testWorkbook.getCachedObjectCount();
            ctx.sync().then(function () {
                RichApiTest.log.comment("CountEnd=" + countEnd.value);
                if (countEnd.value != countStart.value) {
                    success = false;
                }
                RichApiTest.log.done(success);
            }, FakeExcelTest.reportError);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testKeepReferenceNoLoad = testKeepReferenceNoLoad;
    // $top, $skip test.
    // For charts, we use getItemAt for enumeration. For Worksheets, we use _NewEnum for enumeration
    function testChartsTop() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var charts = ctx.application.activeWorkbook.charts;
        ctx.load(charts, { select: "Name, chartType", top: 2 });
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("Count=" + charts.count);
            RichApiTest.log.comment("The length=" + charts.items.length);
            if (charts.items.length != 2) {
                RichApiTest.log.comment("The length is not 2");
                success = false;
            }
            for (var i = 0; i < charts.items.length; i++) {
                RichApiTest.log.comment("Name=" + charts.items[i].name);
                RichApiTest.log.comment("ChartType=" + charts.items[i].chartType);
            }
            jQuery.ajax(OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/charts?$top=2").done(function (data) {
                RichApiTest.log.comment(data);
                var v = JSON.parse(data);
                if (v.value.length != 2) {
                    RichApiTest.log.comment("The rest length is not 2");
                    success = false;
                }
                RichApiTest.log.done(success);
            }).fail(FakeExcelTest.reportAjaxError);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testChartsTop = testChartsTop;
    function testChartsTopSkip() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var charts = ctx.application.activeWorkbook.charts;
        ctx.load(charts, { select: "Name, chartType", top: 2, skip: 2 });
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("Count=" + charts.count);
            RichApiTest.log.comment("The length=" + charts.items.length);
            if (charts.items.length != 2) {
                RichApiTest.log.comment("The length is not 2");
                success = false;
            }
            for (var i = 0; i < charts.items.length; i++) {
                RichApiTest.log.comment("Name=" + charts.items[i].name);
                RichApiTest.log.comment("ChartType=" + charts.items[i].chartType);
            }
            jQuery.ajax(OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/charts?$top=2&$skip=2").done(function (data) {
                RichApiTest.log.comment(data);
                var v = JSON.parse(data);
                if (v.value.length != 2) {
                    RichApiTest.log.comment("The rest length is not 2");
                    success = false;
                }
                RichApiTest.log.done(success);
            }).fail(FakeExcelTest.reportAjaxError);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testChartsTopSkip = testChartsTopSkip;
    // skip to almost the end
    function testChartsTopSkipEnd() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var charts = ctx.application.activeWorkbook.charts;
        ctx.load(charts);
        ctx.sync().then(function () {
            var count = charts.count;
            var skip = count - 2;
            ctx.load(charts, { 'top': 20, 'skip': skip });
            ctx.sync().then(function () {
                var success = true;
                RichApiTest.log.comment("Count=" + charts.count);
                RichApiTest.log.comment("The length=" + charts.items.length);
                if (charts.items.length != 2) {
                    RichApiTest.log.comment("The length is not 2");
                    success = false;
                }
                for (var i = 0; i < charts.items.length; i++) {
                    RichApiTest.log.comment("Name=" + charts.items[i].name);
                    RichApiTest.log.comment("ChartType=" + charts.items[i].chartType);
                }
                jQuery.ajax(OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/charts?$top=20&$skip=" + skip).done(function (data) {
                    RichApiTest.log.comment(data);
                    var v = JSON.parse(data);
                    if (v.value.length != 2) {
                        RichApiTest.log.comment("The rest length is not 2");
                        success = false;
                    }
                    RichApiTest.log.done(success);
                }).fail(FakeExcelTest.reportAjaxError);
            }, FakeExcelTest.reportError);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testChartsTopSkipEnd = testChartsTopSkipEnd;
    function testWorksheetsTop() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var sheets = ctx.application.activeWorkbook.sheets;
        ctx.load(sheets, { top: 2 });
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("The length=" + sheets.items.length);
            if (sheets.items.length != 2) {
                RichApiTest.log.comment("The length is not 2");
                success = false;
            }
            for (var i = 0; i < sheets.items.length; i++) {
                RichApiTest.log.comment("Name=" + sheets.items[i].name);
                RichApiTest.log.comment("_Id=" + sheets.items[i]._Id);
                success = success && verifySheet(sheets.items[i], i);
            }
            jQuery.ajax(OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/sheets?$top=2").done(function (data) {
                RichApiTest.log.comment(data);
                var v = JSON.parse(data);
                if (v.value.length != 2) {
                    RichApiTest.log.comment("The rest length is not 2");
                    success = false;
                }
                RichApiTest.log.done(success);
            }).fail(FakeExcelTest.reportAjaxError);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testWorksheetsTop = testWorksheetsTop;
    function testWorksheetsTopSkip() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var sheets = ctx.application.activeWorkbook.sheets;
        ctx.load(sheets, { top: 2, skip: 2 });
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("The length=" + sheets.items.length);
            if (sheets.items.length != 2) {
                RichApiTest.log.comment("The length is not 2");
                success = false;
            }
            for (var i = 0; i < sheets.items.length; i++) {
                RichApiTest.log.comment("Name=" + sheets.items[i].name);
                RichApiTest.log.comment("_Id=" + sheets.items[i]._Id);
                success = success && verifySheet(sheets.items[i], i + 2);
            }
            jQuery.ajax(OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/sheets?$top=2&$skip=2").done(function (data) {
                RichApiTest.log.comment(data);
                var v = JSON.parse(data);
                if (v.value.length != 2) {
                    RichApiTest.log.comment("The rest length is not 2");
                    success = false;
                }
                RichApiTest.log.done(success);
            }).fail(FakeExcelTest.reportAjaxError);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testWorksheetsTopSkip = testWorksheetsTopSkip;
    function testWorksheetsTopSkipEnd() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var sheets = ctx.application.activeWorkbook.sheets;
        ctx.load(sheets);
        ctx.sync().then(function () {
            var count = sheets.items.length;
            RichApiTest.log.comment("Count=" + count);
            var skip = count - 2;
            ctx.load(sheets, { 'skip': skip, top: 20 });
            ctx.sync().then(function () {
                var success = true;
                RichApiTest.log.comment("The length=" + sheets.items.length);
                if (sheets.items.length != 2) {
                    RichApiTest.log.comment("The length is not 2");
                    success = false;
                }
                for (var i = 0; i < sheets.items.length; i++) {
                    RichApiTest.log.comment("Name=" + sheets.items[i].name);
                    RichApiTest.log.comment("_Id=" + sheets.items[i]._Id);
                }
                jQuery.ajax(OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/sheets?$top=2&$skip=" + skip).done(function (data) {
                    RichApiTest.log.comment(data);
                    var v = JSON.parse(data);
                    if (v.value.length != 2) {
                        RichApiTest.log.comment("The rest length is not 2");
                        success = false;
                    }
                    RichApiTest.log.done(success);
                }).fail(FakeExcelTest.reportAjaxError);
            }, FakeExcelTest.reportError);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testWorksheetsTopSkipEnd = testWorksheetsTopSkipEnd;
    function testExpand() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var book = ctx.application.activeWorkbook;
        ctx.load(book, { expand: "sheEts, charTs" });
        ctx.sync().then(function () {
            var count = book.sheets.items.length;
            RichApiTest.log.comment("SheetsCount=" + count);
            for (var i = 0; i < book.sheets.items.length; i++) {
                RichApiTest.log.comment("SheetName=" + book.sheets.items[i].name);
            }
            count = book.charts.items.length;
            RichApiTest.log.comment("ChartsCount=" + count);
            for (var i = 0; i < book.charts.items.length; i++) {
                RichApiTest.log.comment("SheetName=" + book.charts.items[i].name);
            }
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testExpand = testExpand;
    function testExpand_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook?$expand=" + encodeURIComponent("sheeTs, charTs"),
            method: RichApiTest.RestUtility.httpMethodGet,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            var v = JSON.parse(resp.body);
            var success = true;
            if (v.charts.length < 2) {
                throw "The charts length is less than 2";
            }
            for (var i = 0; i < v.charts.length; i++) {
                var expectedId = "activeWorkbook/charts(" + v.charts[i].id + ")";
                if (v.charts[i]["@odata.id"] != expectedId) {
                    throw "Chart[" + i + "] id is " + v.charts[i]["@odata.id"] + ", not expected " + expectedId;
                }
            }
            if (v.sheets.length < 2) {
                throw "The sheets length is less than 2";
            }
            for (var i = 0; i < v.sheets.length; i++) {
                if (v.sheets[i]._Id) {
                    throw "Should not have _Id";
                }
                var expectedIdPrefix = "activeWorkbook/sheets(";
                var odataid = v.sheets[i]["@odata.id"];
                if (odataid.substr(0, expectedIdPrefix.length) != expectedIdPrefix) {
                    throw "Sheet[" + i + "] id is not expected " + expectedIdPrefix;
                }
            }
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testExpand_rest = testExpand_rest;
    function testSelectExpand() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var book = ctx.application.activeWorkbook;
        ctx.load(book, { select: "sheets/name,charts/name", expand: "sheEts, charTs" });
        ctx.sync().then(function () {
            var count = book.sheets.items.length;
            RichApiTest.log.comment("SheetsCount=" + count);
            for (var i = 0; i < book.sheets.items.length; i++) {
                RichApiTest.log.comment("SheetName=" + book.sheets.items[i].name);
            }
            count = book.charts.items.length;
            RichApiTest.log.comment("ChartsCount=" + count);
            for (var i = 0; i < book.charts.items.length; i++) {
                RichApiTest.log.comment("SheetName=" + book.charts.items[i].name);
            }
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testSelectExpand = testSelectExpand;
    function testSelectExpandWithoutExpand() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var book = ctx.application.activeWorkbook;
        ctx.load(book, { select: "sheets/name,charts/name" });
        ctx.sync().then(function () {
            var count = book.sheets.items.length;
            RichApiTest.log.comment("SheetsCount=" + count);
            for (var i = 0; i < book.sheets.items.length; i++) {
                RichApiTest.log.comment("SheetName=" + book.sheets.items[i].name);
            }
            count = book.charts.items.length;
            RichApiTest.log.comment("ChartsCount=" + count);
            for (var i = 0; i < book.charts.items.length; i++) {
                RichApiTest.log.comment("ChartName=" + book.charts.items[i].name);
            }
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testSelectExpandWithoutExpand = testSelectExpandWithoutExpand;
    function testSelectExpandWithoutExpand2() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var book = ctx.application.activeWorkbook;
        ctx.load(book, { select: ["sheets/name", "charts/name", "charts/title"] });
        ctx.sync().then(function () {
            var count = book.sheets.items.length;
            RichApiTest.log.comment("SheetsCount=" + count);
            for (var i = 0; i < book.sheets.items.length; i++) {
                RichApiTest.log.comment("SheetName=" + book.sheets.items[i].name);
            }
            count = book.charts.items.length;
            RichApiTest.log.comment("ChartsCount=" + count);
            for (var i = 0; i < book.charts.items.length; i++) {
                RichApiTest.log.comment("ChartName=" + book.charts.items[i].name);
                RichApiTest.log.comment("ChartTitle=" + book.charts.items[i].title);
            }
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testSelectExpandWithoutExpand2 = testSelectExpandWithoutExpand2;
    function testSelectExpandError() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var book = ctx.application.activeWorkbook;
        var success = true;
        try {
            ctx.load(book, { select: true });
            RichApiTest.log.comment("does not get exptected exception");
            success = false;
        }
        catch (ex) {
            if (ex.code != "InvalidArgument") {
                RichApiTest.log.comment("Wrong error code: " + ex.code);
                success = false;
            }
        }
        try {
            ctx.load(book, { expand: true });
            RichApiTest.log.comment("does not get exptected exception");
            success = false;
        }
        catch (ex) {
            if (ex.code != "InvalidArgument") {
                RichApiTest.log.comment("Wrong error code: " + ex.code);
                success = false;
            }
        }
        RichApiTest.log.done(success);
    }
    FakeExcelTest.testSelectExpandError = testSelectExpandError;
    function testSelectExpand_rest() {
        jQuery.ajax(OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook?$select=sheets/name&$expand=sheeTs,charTs").done(function (data) {
            RichApiTest.log.comment(data);
            var v = JSON.parse(data);
            var success = true;
            if (v.charts.length < 2) {
                RichApiTest.log.comment("The charts length is less than 2");
                success = false;
            }
            if (v.sheets.length < 2) {
                RichApiTest.log.comment("The sheets length is less than 2");
                success = false;
            }
            RichApiTest.log.done(success);
        }).fail(FakeExcelTest.reportAjaxError);
    }
    FakeExcelTest.testSelectExpand_rest = testSelectExpand_rest;
    function testNotLoaded() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var charts = ctx.application.activeWorkbook.charts.load();
        var success = true;
        try {
            var items = charts.items;
            success = false;
        }
        catch (e) {
            RichApiTest.log.comment("Get expected exception " + e);
        }
        var chart = charts.getItem("Chart1000").load({ select: "name" });
        try {
            RichApiTest.log.comment("Should not get chart.name " + chart.name);
            success = false;
        }
        catch (e) {
            RichApiTest.log.comment("Get expected exception " + e);
        }
        ctx.sync().then(function () {
            for (var i = 0; i < charts.items.length; i++) {
                RichApiTest.log.comment("chart " + charts.items[i].name);
            }
            RichApiTest.log.comment("chart.name = " + chart.name);
            RichApiTest.log.done(success);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testNotLoaded = testNotLoaded;
    function testOnAccessSetProperty() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range1 = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        var range2 = ctx.application.activeWorkbook.activeWorksheet.range("B2");
        range1.text = "A1";
        range2.text = "B2";
        range1.load();
        range2.load();
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("range1.logText=" + range1.logText);
            // verify that the _OnAccess method was called.
            if (range1.logText.indexOf("_OnAccess") < 0) {
                RichApiTest.log.comment("range1.logText does not have _OnAccess");
                success = false;
            }
            RichApiTest.log.comment("range2.logText=" + range2.logText);
            if (range2.logText.indexOf("_OnAccess") < 0) {
                RichApiTest.log.comment("range2.logText does not have _OnAccess");
                success = false;
            }
            RichApiTest.log.done(success);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testOnAccessSetProperty = testOnAccessSetProperty;
    function testOnAccessLoad() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range1 = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        var range2 = ctx.application.activeWorkbook.activeWorksheet.range("B2");
        range1.load();
        range2.load();
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("range1.logText=" + range1.logText);
            // verify that the _OnAccess method was called.
            if (range1.logText.indexOf("_OnAccess") < 0) {
                RichApiTest.log.comment("range1.logText does not have _OnAccess");
                success = false;
            }
            RichApiTest.log.comment("range2.logText=" + range2.logText);
            if (range2.logText.indexOf("_OnAccess") < 0) {
                RichApiTest.log.comment("range2.logText does not have _OnAccess");
                success = false;
            }
            RichApiTest.log.done(success);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testOnAccessLoad = testOnAccessLoad;
    function testOnAccessMethod() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range1 = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        var range2 = ctx.application.activeWorkbook.activeWorksheet.range("B2");
        range1.activate();
        range2.activate();
        range1.load();
        range2.load();
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("range1.logText=" + range1.logText);
            // verify that the _OnAccess method was called.
            if (range1.logText.indexOf("_OnAccess") < 0) {
                RichApiTest.log.comment("range1.logText does not have _OnAccess");
                success = false;
            }
            RichApiTest.log.comment("range2.logText=" + range2.logText);
            if (range2.logText.indexOf("_OnAccess") < 0) {
                RichApiTest.log.comment("range2.logText does not have _OnAccess");
                success = false;
            }
            RichApiTest.log.done(success);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testOnAccessMethod = testOnAccessMethod;
    function testChartsForEach() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var charts = ctx.application.activeWorkbook.charts;
        forEachAsync(charts, function (chart) {
            // update the chart name to be something else
            RichApiTest.log.comment("Chart name=" + chart.name);
            chart.name = "New" + chart.name;
        }, function () {
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testChartsForEach = testChartsForEach;
    function forEachAsync(charts, action, successCallback, failureCallback) {
        forEachAsyncHelper(charts, 3, 0, action, successCallback, failureCallback);
    }
    function forEachAsyncHelper(charts, top, skip, action, successCallback, failureCallback) {
        var ctx = charts.context;
        ctx.load(charts, { top: top, skip: skip });
        ctx.sync().then(function () {
            for (var i = 0; i < charts.items.length; ++i) {
                action(charts.items[i]);
            }
            if (charts.items.length == 0) {
                successCallback();
            }
            else {
                forEachAsyncHelper(charts, top, skip + charts.items.length, action, successCallback, failureCallback);
            }
        }).catch(function (error) {
            failureCallback(error);
        });
    }
    function testChartsForEachPageRead() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var charts = ctx.application.activeWorkbook.charts;
        var pageSize = 2;
        forEachPageAsync(charts, pageSize, function (items) {
            for (var i = 0; i < items.length; i++) {
                RichApiTest.log.comment(items[i].name);
            }
            return null;
        }).then(function () {
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.testChartsForEachPageRead = testChartsForEachPageRead;
    function testChartsForEachPageReadWrite() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var charts = ctx.application.activeWorkbook.charts;
        var pageSize = 2;
        forEachPageAsync(charts, pageSize, function (items) {
            for (var i = 0; i < items.length; i++) {
                RichApiTest.log.comment(items[i].name);
            }
            for (var i = 0; i < items.length; i++) {
                items[i].title = "New Title";
            }
            // executeQuery to save the change for the page
            return ctx.sync();
        }).then(function () {
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.testChartsForEachPageReadWrite = testChartsForEachPageReadWrite;
    function forEachPageAsync(collection, pageSize, action) {
        return forEachPageAsyncOnePage(collection, pageSize, 0, action);
    }
    FakeExcelTest.forEachPageAsync = forEachPageAsync;
    function forEachPageAsyncOnePage(collection, top, skip, action) {
        var ctx = collection.context;
        ctx.load(collection, { top: top, skip: skip });
        RichApiTest.log.comment("Fetching page top=" + top + ", skip=" + skip);
        return ctx.sync().then(function () {
            if (collection.items.length == top) {
                // there could be more and we need to load another page after action.
                return OfficeExtension.Utility._createPromiseFromResult(null).then(function () {
                    return action(collection.items);
                }).then(function () {
                    return forEachPageAsyncOnePage(collection, top, skip + top, action);
                });
            }
            else {
                // there is no more page and we only need to trigger action.
                return action(collection.items);
            }
        });
    }
    FakeExcelTest.forEachPageAsyncOnePage = forEachPageAsyncOnePage;
    function testParamValidationBool() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var testCase = FakeExcelApi.TestCaseObject.newObject(ctx);
        var boolNull = testCase.testParamBool(null);
        var boolTrue = testCase.testParamBool(true);
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("boolNull=" + boolNull.value);
            if (boolNull.value !== false) {
                RichApiTest.log.comment("null is not coverted to false");
                success = false;
            }
            RichApiTest.log.comment("boolTrue=" + boolTrue.value);
            if (boolTrue.value !== true) {
                RichApiTest.log.comment("true is not true");
                success = false;
            }
            testCase.testParamBool(("true"));
            ctx.sync().then(function () {
                RichApiTest.log.comment("Should get failure");
                success = false;
                RichApiTest.log.done(success);
            }, function (result) {
                RichApiTest.log.done(success && result.code == FakeExcelApi.ErrorCodes.invalidArgument);
            });
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testParamValidationBool = testParamValidationBool;
    function testParamValidationInt() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var testCase = FakeExcelApi.TestCaseObject.newObject(ctx);
        var resultNull = testCase.testParamInt(null);
        var resultNull2 = testCase.testParamInt();
        var resultInteger = testCase.testParamInt(3);
        var resultDouble = testCase.testParamInt(3.0);
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("resultNull=" + resultNull.value);
            if (resultNull.value !== 0) {
                RichApiTest.log.comment("null is not coverted to 0");
                success = false;
            }
            RichApiTest.log.comment("resultNull2=" + resultNull2.value);
            if (resultNull2.value !== 0) {
                RichApiTest.log.comment("null2 is not coverted to 0");
                success = false;
            }
            RichApiTest.log.comment("resultInteger=" + resultInteger.value);
            if (resultInteger.value !== 3) {
                RichApiTest.log.comment("3 is not 3");
                success = false;
            }
            RichApiTest.log.comment("resultDouble=" + resultDouble.value);
            if (resultDouble.value !== 3) {
                RichApiTest.log.comment("3.0 is not 3");
                success = false;
            }
            testCase.testParamInt(3.5);
            ctx.sync().then(function () {
                RichApiTest.log.comment("Should get failure as we pass 3.5 as integer value");
                success = false;
                RichApiTest.log.done(success);
            }, function (result) {
                RichApiTest.log.done(success && result.code == FakeExcelApi.ErrorCodes.invalidArgument);
            });
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testParamValidationInt = testParamValidationInt;
    function testParamValidationDouble() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var testCase = FakeExcelApi.TestCaseObject.newObject(ctx);
        var resultNull = testCase.testParamDouble(null);
        var resultInteger = testCase.testParamDouble(3);
        var resultDouble = testCase.testParamDouble(3.0);
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("resultNull=" + resultNull.value);
            if (resultNull.value !== 0.0) {
                RichApiTest.log.comment("null is not coverted to 0");
                success = false;
            }
            RichApiTest.log.comment("resultInteger=" + resultInteger.value);
            if (resultInteger.value !== 3.0) {
                RichApiTest.log.comment("3 is not 3.0");
                success = false;
            }
            RichApiTest.log.comment("resultDouble=" + resultDouble.value);
            if (resultDouble.value !== 3.0) {
                RichApiTest.log.comment("3.0 is not 3.0");
                success = false;
            }
            testCase.testParamDouble("3.5");
            ctx.sync().then(function () {
                RichApiTest.log.comment("Should get failure as we use string 3.5 for double");
                success = false;
                RichApiTest.log.done(success);
            }, function (result) {
                RichApiTest.log.done(success && result.code == FakeExcelApi.ErrorCodes.invalidArgument);
            });
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testParamValidationDouble = testParamValidationDouble;
    function testParamValidationDouble2() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var testCase = FakeExcelApi.TestCaseObject.newObject(ctx);
        var inputs = [
            Math.sin(0.00000000001),
            Math.sin(0.0000001),
            Math.sin(0.01),
            Math.sin(1),
            0.2,
            123 + Math.sin(1),
            1234567 + Math.sin(1),
            12345678901234 + Math.sin(1),
            123456789012346789 + Math.sin(1)
        ];
        var results = [];
        for (var i = 0; i < inputs.length; i++) {
            var result = testCase.testParamDouble(inputs[i]);
            results.push(result);
        }
        ctx.sync().then(function () {
            for (var i = 0; i < inputs.length; i++) {
                RichApiTest.log.comment("input =" + inputs[i]);
                RichApiTest.log.comment("result=" + results[i].value);
                if (results[i].value !== inputs[i]) {
                    RichApiTest.log.comment("Not same");
                }
            }
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testParamValidationDouble2 = testParamValidationDouble2;
    function testParamValidationString() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var testCase = FakeExcelApi.TestCaseObject.newObject(ctx);
        var resultNull = testCase.testParamString();
        var resultNull2 = testCase.testParamString(null);
        var resultString = testCase.testParamString("abc");
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("resultNull=" + resultNull.value);
            if (resultNull.value !== null) {
                RichApiTest.log.comment("null is not coverted to null");
                success = false;
            }
            RichApiTest.log.comment("resultNull2=" + resultNull2.value);
            if (resultNull2.value !== null) {
                RichApiTest.log.comment("null2 is not coverted to null");
                success = false;
            }
            RichApiTest.log.comment("resultString=" + resultString.value);
            if (resultString.value !== "abc") {
                RichApiTest.log.comment("abc is not abc");
                success = false;
            }
            testCase.testParamString(3);
            ctx.sync().then(function () {
                RichApiTest.log.comment("Should get failure as we use integer 3 for string");
                success = false;
                RichApiTest.log.done(success);
            }, function (result) {
                RichApiTest.log.done(success && result.code == FakeExcelApi.ErrorCodes.invalidArgument);
            });
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testParamValidationString = testParamValidationString;
    function testParamValidationRange() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var testCase = FakeExcelApi.TestCaseObject.newObject(ctx);
        var range = ctx.application.activeWorkbook.activeWorksheet.range("C1");
        var resultNull = testCase.testParamRange(null);
        var resultRange = testCase.testParamRange(range);
        ctx.sync().then(function () {
            var success = true;
            testCase.testParamRange(3);
            ctx.sync().then(function () {
                RichApiTest.log.comment("Should get failure");
                success = false;
                RichApiTest.log.done(success);
            }, function (result) {
                RichApiTest.log.done(success && result.code == FakeExcelApi.ErrorCodes.invalidArgument);
            });
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testParamValidationRange = testParamValidationRange;
    function testObjectPathExp() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var testCase = FakeExcelApi.TestCaseObject.newObject(ctx);
        var exp;
        var success = true;
        exp = OfficeExtension.Utility.getObjectPathExpression(testCase._objectPath);
        RichApiTest.log.comment("testCase=" + exp);
        if (exp != "new()") {
            RichApiTest.log.comment("!= new()");
            success = false;
        }
        exp = OfficeExtension.Utility.getObjectPathExpression(ctx.application._objectPath);
        RichApiTest.log.comment("application=" + exp);
        if (exp != "") {
            RichApiTest.log.comment("!= empty");
            success = false;
        }
        exp = OfficeExtension.Utility.getObjectPathExpression(ctx.application.activeWorkbook.sheets.getItem('Sheet')._objectPath);
        RichApiTest.log.comment("sheet=" + exp);
        if (exp != "activeWorkbook.sheets.getItem()") {
            RichApiTest.log.comment("!= activeWorkbook.sheets.getItem()");
            success = false;
        }
        exp = OfficeExtension.Utility.getObjectPathExpression(ctx.application.activeWorkbook.sheets.getItem('Sheet').range('A1')._objectPath);
        RichApiTest.log.comment("range=" + exp);
        if (exp != "activeWorkbook.sheets.getItem().range()") {
            RichApiTest.log.comment("!= activeWorkbook.sheets.getItem().range()");
            success = false;
        }
        RichApiTest.log.done(success);
    }
    FakeExcelTest.testObjectPathExp = testObjectPathExp;
    function testObjectPathInvalid() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range = ctx.application.activeWorkbook.activeWorksheet.getActiveCellInvalidAfterRequest();
        ctx.sync().then(function () {
            var success = true;
            try {
                ctx.load(range);
                success = false;
            }
            catch (ex) {
                var msg = ex.toString();
                RichApiTest.log.comment("Get expected exception: " + ex);
            }
            RichApiTest.log.done(success);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testObjectPathInvalid = testObjectPathInvalid;
    function testInvalidDispatchParameter() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range = ctx.application.testWorkbook;
        var result = ctx.application.activeWorkbook.activeWorksheet.someRangeOperation("Abc", range);
        ctx.sync().then(function () {
            RichApiTest.log.fail("Should fail");
        }, function (errorInfo) {
            RichApiTest.log.comment("Get exptected failure");
            RichApiTest.log.comment(errorInfo.code);
            RichApiTest.log.comment(errorInfo.message);
            RichApiTest.log.done(true);
        });
    }
    FakeExcelTest.testInvalidDispatchParameter = testInvalidDispatchParameter;
    function testInvalidClientContext() {
        var ctx1 = new FakeExcelApi.ExcelClientRequestContext();
        var range = ctx1.application.activeWorkbook.activeWorksheet.getActiveCell();
        var ctx2 = new FakeExcelApi.ExcelClientRequestContext();
        var success = true;
        try {
            ctx2.load(range);
            RichApiTest.log.comment("Should not be able to load range");
            success = false;
        }
        catch (ex) {
            RichApiTest.log.comment("Get expected ex " + ex);
        }
        try {
            ctx2.application.activeWorkbook.activeWorksheet.someRangeOperation("a", range);
            RichApiTest.log.comment("Should not be able to use range");
            success = false;
        }
        catch (ex) {
            RichApiTest.log.comment("Get expected ex " + ex);
        }
        RichApiTest.log.done(success);
    }
    FakeExcelTest.testInvalidClientContext = testInvalidClientContext;
    function testPromisesExecuteAsync() {
        // This function tests that both the standard "return ctx.sync()" (within a lambda) and
        // the standalone ".then(ctx.sync)" (outside a lambda) work.
        // The latter used to fail due to a bug described in OfficeMain:2382501
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        ctx.trackedObjects.add(range);
        var result1 = range.replaceValue("Hello");
        ctx.load(range, "Value, RowIndex");
        ctx.sync().then(function () {
            RichApiTest.log.comment("Initial executeAsync");
            range.replaceValue("Hello2");
            return ctx.sync();
        }).then(function () {
            RichApiTest.log.comment('Successfully completed a "return ctx.sync()"');
            range.replaceValue("Hello2");
        }).then(ctx.sync).then(function () {
            RichApiTest.log.comment('Successfully completed a ".then(ctx.sync)"');
            RichApiTest.log.pass("Both syntaxes completed successfully");
        }).catch(function (e) {
            RichApiTest.log.comment(JSON.stringify(e));
        });
    }
    FakeExcelTest.testPromisesExecuteAsync = testPromisesExecuteAsync;
    function testEnumArray() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range = ctx.application.activeWorkbook.activeWorksheet.getActiveCell();
        ctx.load(range);
        ctx.sync().then(function () {
            var success = true;
            for (var i = 0; i < range.valueTypes.length; i++) {
                if (range.valueTypes[i][0] != 'Unknown') {
                    success = false;
                }
                if (range.valueTypes[i][1] != 'Empty') {
                    success = false;
                }
                if (range.valueTypes[i][2] != 'String') {
                    success = false;
                }
                for (var j = 0; j < range.valueTypes[i].length; j++) {
                    RichApiTest.log.comment("range.valueTypes[" + i + "][" + j + "]=" + range.valueTypes[i][j]);
                }
            }
            RichApiTest.log.done(success);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testEnumArray = testEnumArray;
    function testWacFind() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var sheet = ctx.application.activeWorkbook.sheets.findSheet("foo");
        var range1 = sheet.range("A1");
        var result1 = sheet.someRangeOperation("Hello", range1);
        var range2 = sheet.range("A2");
        var result2 = sheet.someRangeOperation("World", range2);
        ctx.load(range1);
        ctx.load(range2);
        ctx.sync().then(function () {
            var success = true;
            RichApiTest.log.comment("Result1=" + result1.value);
            if (result1.value != "Hello") {
                RichApiTest.log.comment("Result1.value != Hello");
                success = false;
            }
            RichApiTest.log.comment("Range1.text=" + range1.text);
            if (range1.text != "Hello") {
                RichApiTest.log.comment("range1.text != Hello");
                success = false;
            }
            RichApiTest.log.comment("Result2=" + result2.value);
            if (result2.value != "World") {
                RichApiTest.log.comment("Result2.value != World");
                success = false;
            }
            RichApiTest.log.comment("Range2.text=" + range2.text);
            if (range2.text != "World") {
                RichApiTest.log.comment("range2.text != World");
                success = false;
            }
            RichApiTest.log.done(success);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testWacFind = testWacFind;
    function testRequestFlagHack1() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var sheet = ctx.application.activeWorkbook.sheets.findSheet("foo");
        var range1 = sheet.range("A1");
        range1.value = "test";
        (ctx._pendingRequest).m_flags = 0 /* None */;
        ctx.sync().then(function () {
            RichApiTest.log.comment("Success but should fail");
            RichApiTest.log.done(false);
        }, function (err) {
            FakeExcelTest.logError(err);
            RichApiTest.log.done(err.code === FakeExcelApi.ErrorCodes.accessDenied);
        });
    }
    FakeExcelTest.testRequestFlagHack1 = testRequestFlagHack1;
    function testRequestFlagHack2() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var sheet = ctx.application.activeWorkbook.sheets.findSheet("foo");
        var range1 = sheet.range("A1");
        sheet.someRangeOperation("text", range1);
        (ctx._pendingRequest).m_flags = 0 /* None */;
        ctx.sync().then(function () {
            RichApiTest.log.comment("Success but should fail");
            RichApiTest.log.done(false);
        }, function (err) {
            FakeExcelTest.logError(err);
            RichApiTest.log.done(err.code === FakeExcelApi.ErrorCodes.accessDenied);
        });
    }
    FakeExcelTest.testRequestFlagHack2 = testRequestFlagHack2;
    function testNullRange() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var sheet = ctx.application.activeWorkbook.activeWorksheet;
        var range1 = sheet.nullRange("A1");
        var range2 = sheet.nullRange("A2");
        ctx.sync().then(function () {
            RichApiTest.log.done(true);
        }, FakeExcelTest.reportError);
    }
    FakeExcelTest.testNullRange = testNullRange;
    function testRange_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/activeWorksheet/Range('A1')",
            method: RichApiTest.RestUtility.httpMethodGet
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            request.url = request.url + "/text";
            return RichApiTest.RestUtility.invoke(request);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testRange_rest = testRange_rest;
    function testNullRangeGet_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/activeWorksheet/nullRange('A1')",
            method: RichApiTest.RestUtility.httpMethodGet
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusNotFound);
            request.url = request.url + "/text";
            return RichApiTest.RestUtility.invoke(request);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusBadRequest);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testNullRangeGet_rest = testNullRangeGet_rest;
    function testNullRangePatch_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/activeWorksheet/nullRange('A1')",
            method: RichApiTest.RestUtility.httpMethodPatch,
            body: "{\"value\": \"abc\"}",
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusNotFound);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testNullRangePatch_rest = testNullRangePatch_rest;
    function testNullRangeDelete_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/activeWorksheet/nullRange('A1')",
            method: RichApiTest.RestUtility.httpMethodDelete,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusMethodNotAllowed);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testNullRangeDelete_rest = testNullRangeDelete_rest;
    function testNullChartDelete_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/activeWorksheet/nullChart('A1')",
            method: RichApiTest.RestUtility.httpMethodDelete,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusNoContent);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testNullChartDelete_rest = testNullChartDelete_rest;
    function testVoidMethod_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/activeWorksheet/range('A1')/activate",
            method: RichApiTest.RestUtility.httpMethodPost,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusNoContent);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testVoidMethod_rest = testVoidMethod_rest;
    function testScalarMethod_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/activeWorksheet/range('A1')/replacevalue('abc')",
            method: RichApiTest.RestUtility.httpMethodPost,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            var result = JSON.parse(resp.body);
            if (result["value"] != "Initial Value") {
                throw "Not get expected Initial Value";
            }
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testScalarMethod_rest = testScalarMethod_rest;
    function testBlockedMethodKeepReference_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/activeWorksheet/range('A1')/_KeepReference",
            method: RichApiTest.RestUtility.httpMethodPost,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusBadRequest);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testBlockedMethodKeepReference_rest = testBlockedMethodKeepReference_rest;
    function testBlockedMethodNotRest_rest() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var result = ctx.application.activeWorkbook.activeWorksheet.range("A1").notRestMethod();
        ctx.sync().then(function () {
            if (result.value != 1234) {
                throw "Not Rest method failed";
            }
        }).then(function () {
            var request = {
                url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/activeWorksheet/range('A1')/NotRestMethod",
                method: RichApiTest.RestUtility.httpMethodGet,
            };
            return RichApiTest.RestUtility.invoke(request);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusBadRequest);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testBlockedMethodNotRest_rest = testBlockedMethodNotRest_rest;
    function testBlockedMethodIndexer_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/sheets/.indexer('Sheet1')",
            method: RichApiTest.RestUtility.httpMethodGet,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusBadRequest);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testBlockedMethodIndexer_rest = testBlockedMethodIndexer_rest;
    function testBlockedProperty_rest() {
        var sheetUrl = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/sheets('Sheet2001')";
        var request = {
            url: sheetUrl,
            method: RichApiTest.RestUtility.httpMethodGet,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            request.url = sheetUrl + "/name";
            return RichApiTest.RestUtility.invoke(request);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            request.url = sheetUrl + "/_Id";
            return RichApiTest.RestUtility.invoke(request);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusBadRequest);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testBlockedProperty_rest = testBlockedProperty_rest;
    function testExcludedProperty_rest() {
        var sheetUrl = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/sheets('Sheet2001')";
        var request = {
            url: sheetUrl,
            method: RichApiTest.RestUtility.httpMethodGet,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            var sheet = JSON.parse(resp.body);
            if (sheet.calculatedName) {
                throw "CalculatedName is supposed to be excluded";
            }
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testExcludedProperty_rest = testExcludedProperty_rest;
    function testExcludedProperty2_rest() {
        var bookUrl = OfficeExtension.Constants.localDocumentApiPrefix + "testWorkbook";
        var request = {
            url: bookUrl,
            method: RichApiTest.RestUtility.httpMethodGet,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
        }).then(function () {
            // the ErrorWorksheet was marked as ExcludedFromRest
            request.url = bookUrl + "/ErrorWorksheet";
            return RichApiTest.RestUtility.invoke(request);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusBadRequest);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testExcludedProperty2_rest = testExcludedProperty2_rest;
    function testPostToPrimitiveProperty_rest() {
        var sheetUrl = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/sheets('Sheet2001')";
        var request = {
            url: sheetUrl,
            method: RichApiTest.RestUtility.httpMethodGet,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            request.url = sheetUrl + "/name";
            return RichApiTest.RestUtility.invoke(request);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            request.method = RichApiTest.RestUtility.httpMethodPost;
            request.url = sheetUrl + "/name";
            return RichApiTest.RestUtility.invoke(request);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusBadRequest);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testPostToPrimitiveProperty_rest = testPostToPrimitiveProperty_rest;
    function testPatchToPrimitiveProperty_rest() {
        var sheetUrl = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/sheets('Sheet2001')";
        var request = {
            url: sheetUrl,
            method: RichApiTest.RestUtility.httpMethodGet,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            request.url = sheetUrl + "/name";
            return RichApiTest.RestUtility.invoke(request);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            request.method = RichApiTest.RestUtility.httpMethodPatch;
            request.url = sheetUrl + "/name";
            return RichApiTest.RestUtility.invoke(request);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusBadRequest);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testPatchToPrimitiveProperty_rest = testPatchToPrimitiveProperty_rest;
    function testPostToObjectProperty_rest() {
        var sheetUrl = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/sheets('Sheet2001')";
        var request = {
            url: sheetUrl,
            method: RichApiTest.RestUtility.httpMethodGet,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            request.url = sheetUrl + "/ActiveCell";
            return RichApiTest.RestUtility.invoke(request);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            request.method = RichApiTest.RestUtility.httpMethodPost;
            request.url = sheetUrl + "/ActiveCell";
            return RichApiTest.RestUtility.invoke(request);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusBadRequest);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testPostToObjectProperty_rest = testPostToObjectProperty_rest;
    function testPostToFunction_rest() {
        var sheetUrl = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/sheets('Sheet2001')";
        var request = {
            url: sheetUrl,
            method: RichApiTest.RestUtility.httpMethodGet,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            request.url = sheetUrl + "/Range('A1')";
            return RichApiTest.RestUtility.invoke(request);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            request.method = RichApiTest.RestUtility.httpMethodPost;
            request.url = sheetUrl + "/Range('A1')";
            return RichApiTest.RestUtility.invoke(request);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusBadRequest);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testPostToFunction_rest = testPostToFunction_rest;
    function testMethodParameter_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/SomeAction(intVal=1, strVal='abc', enumVal='pie')",
            method: RichApiTest.RestUtility.httpMethodPost,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testMethodParameter_rest = testMethodParameter_rest;
    function testMethodParameterInvalid_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/SomeAction(1, 'abc', 'pie')",
            method: RichApiTest.RestUtility.httpMethodPost,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusBadRequest);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testMethodParameterInvalid_rest = testMethodParameterInvalid_rest;
    function testExecuteAsyncShortCircuits() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var r = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        ctx.trackedObjects.add(r);
        r.text = "New Text";
        var counterAfterFirstExecuteAsync = 0;
        ctx.sync().then(function () {
            // Instrument the "ctx._requestExecutor" to throw an error when it is next invoked.
            // Do it here, rather than initially, because initially it may have been null -- the requestExecutor is only created when needed.
            var originalExecuteAsync = ctx._requestExecutor.executeAsync;
            ctx._requestExecutor.executeAsync = function () {
                counterAfterFirstExecuteAsync++;
                originalExecuteAsync.apply(this, arguments);
            };
            RichApiTest.log.comment("Do nothing that requires an async call");
        }).then(ctx.sync).then(function () {
            if (counterAfterFirstExecuteAsync > 0) {
                throw new Error("Counter of server roundtrips should have remained at 0, whereas it is " + counterAfterFirstExecuteAsync);
            }
        }).then(function () {
            r.text = "Different text";
            ctx.load(r, "text");
        }).then(ctx.sync).then(function () {
            if (counterAfterFirstExecuteAsync != 1) {
                throw new Error("Counter of server roundtrips should now be at 1, since did have a server rountrip (to set text)");
            }
            if (r.text != "Different text") {
                throw new Error("executeAsync failed to apply, range text does not match what it was just set to");
            }
            RichApiTest.log.pass("Successfully short-circuited the empty executeAsync, then successfully applied executeAsync");
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.testExecuteAsyncShortCircuits = testExecuteAsyncShortCircuits;
    function testExecuteAsyncAllowsValuePassThrough1() {
        // Pass through via a return function value
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        ctx.sync().then(function () {
            var r = ctx.application.activeWorkbook.activeWorksheet.range("A1");
            ctx.trackedObjects.add(r);
            r.text = "New Text";
            ctx.load(r, "text");
            return r;
        }).then(ctx.sync).then(function (r) {
            if (r.text != "New Text") {
                throw new Error("Object was not correctly passed through");
            }
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.testExecuteAsyncAllowsValuePassThrough1 = testExecuteAsyncAllowsValuePassThrough1;
    function testExecuteAsyncAllowsValuePassThrough2() {
        // Pass through via a return function value, and with multiple parameters in one object
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        ctx.sync().then(function () {
            var r = ctx.application.activeWorkbook.activeWorksheet.range("A1");
            ctx.trackedObjects.add(r);
            r.text = "New Text";
            ctx.load(r, "text");
            return { range: r, someOtherObject: 5 };
        }).then(ctx.sync).then(function (previousValues) {
            if (previousValues.range.text !== "New Text" || previousValues.someOtherObject !== 5) {
                throw new Error("Object was not correctly passed through");
            }
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.testExecuteAsyncAllowsValuePassThrough2 = testExecuteAsyncAllowsValuePassThrough2;
    function testExecuteAsyncAllowsValuePassThrough3() {
        // Pass through via the "ctx.sync(passThroughObj)" syntax
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        ctx.sync().then(function () {
            var r = ctx.application.activeWorkbook.activeWorksheet.range("A1");
            ctx.trackedObjects.add(r);
            r.text = "New Text";
            ctx.load(r, "text");
            return ctx.sync(r);
        }).then(function (r) {
            if (r.text != "New Text") {
                throw new Error("Object was not correctly passed through");
            }
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.testExecuteAsyncAllowsValuePassThrough3 = testExecuteAsyncAllowsValuePassThrough3;
    function testUrlPathEncode() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var testCase = FakeExcelApi.TestCaseObject.newObject(ctx);
        var values = [
            "\u4e2d\u56fd",
            "\u4e2d\u56fd\uD852\uDF62",
            "http://server/\u4e2d\u56fd",
            "http://server/\u4e2d\u56fdabc",
            "http://server/\u4e2d\u56fdabc?a=b",
            "a b/c",
            "a,b{c}",
        ];
        var value1 = "";
        var value2 = "";
        var value3 = "";
        for (var i = 1; i < 300; i++) {
            var s = String.fromCharCode(i);
            // ? is query string separator
            // the SPHttpUtility.UrlPathEncode also encoded #, &, ' which are not encoded by encodeURI
            if (s != "?" && s != "#" && s != "&" && s != "'") {
                value1 = value1 + s + s;
                value2 = s + value2;
                value3 = value3 + s;
            }
        }
        value3 = value3 + value3;
        values.push("http://server/" + value1);
        values.push("http://server/" + value2);
        values.push("http://server/" + value3);
        var results = [];
        for (var i = 0; i < values.length; i++) {
            results.push(testCase.testUrlPathEncode(values[i]));
        }
        //SPHttpUtility.UrlPathEncode also encoded #, &, ' which are not encoded by encodeURI
        var valueSPHttpUtilitySpecific = "#&'";
        var resultSPHttpUtilitySpecific = testCase.testUrlPathEncode(valueSPHttpUtilitySpecific);
        var encodedSPHttpUtilitySpecific = "%23%26%27";
        ctx.sync().then(function () {
            for (var i = 0; i < values.length; i++) {
                RichApiTest.log.comment("result[" + i + "].value=");
                RichApiTest.log.comment(results[i].value);
                RichApiTest.log.comment("Expected=");
                RichApiTest.log.comment(encodeURI(values[i]));
                if (results[i].value != encodeURI(values[i])) {
                    throw "Expected " + encodeURI(values[i]) + ", actual " + results[i].value;
                }
            }
            RichApiTest.log.comment("resultSPHttpUtilitySpecific.value=");
            RichApiTest.log.comment(resultSPHttpUtilitySpecific.value);
            if (resultSPHttpUtilitySpecific.value != encodedSPHttpUtilitySpecific) {
                throw "Expected " + encodedSPHttpUtilitySpecific + ", actual " + resultSPHttpUtilitySpecific.value;
            }
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.testUrlPathEncode = testUrlPathEncode;
    function testUrlKeyValueDecode() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var testCase = FakeExcelApi.TestCaseObject.newObject(ctx);
        var value1 = "\u4e2d\u56fd";
        RichApiTest.log.comment("Encoded value1=" + encodeURIComponent(value1));
        var result1 = testCase.testUrlKeyValueDecode(encodeURIComponent(value1));
        var value2 = "a,b c";
        RichApiTest.log.comment("Encoded value2=" + encodeURIComponent(value2));
        var result2 = testCase.testUrlKeyValueDecode(encodeURIComponent(value2));
        ctx.sync().then(function () {
            RichApiTest.log.comment("result1.value=" + result1.value);
            RichApiTest.log.comment("result2.value=" + result2.value);
            if (result1.value != value1) {
                throw "Expected " + value1 + ", actual " + result1.value;
            }
            if (result2.value != value2) {
                throw "Expected " + value2 + ", actual " + result2.value;
            }
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.testUrlKeyValueDecode = testUrlKeyValueDecode;
    function test_stream_basic() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var chart = ctx.application.activeWorkbook.getChartByType(FakeExcelApi.ChartType.bar);
        var smallImage = chart.getAsImage(false);
        var largeImage = chart.getAsImage(true);
        ctx.sync().then(function () {
            RichApiTest.log.comment("smallChart.image=");
            RichApiTest.log.comment(smallImage.value);
            RichApiTest.log.comment("largeChart.image=");
            RichApiTest.log.comment(largeImage.value);
            if (smallImage.value != "AAECAwQFBgcICQoLDA0ODxAREhMUFRYXGBkaGxwdHh8gISIjJCUmJygpKissLS4vMDEyMzQ1Njc4OTo7PD0+P0BBQkNERUZHSElKS0xNTk9QUVJTVFVWV1hZWltcXV5fYGFiY2RlZmdoaWprbG1ub3BxcnN0dXZ3eHl6e3x9fn+AgYKDhIWGh4iJiouMjY6PkJGSk5SVlpeYmZqbnJ2en6ChoqOkpaanqKmqq6ytrq+wsbKztLW2t7i5uru8vb6/wMHCw8TFxsfIycrLzM3Oz9DR0tPU1dbX2Nna29zd3t/g4eLj5OXm5+jp6uvs7e7v8PHy8/T19vf4+fr7/P3+/wABAgMEBQYHCAkKCwwNDg8QERITFBUWFxgZGhscHR4fICEiIyQlJicoKSorLC0uLzAxMjM0NTY3ODk6Ozw9Pj9AQUJDREVGR0hJSktMTU5PUFFSU1RVVldYWVpbXF1eX2BhYmNkZWZnaGlqa2xtbm9wcXJzdHV2d3h5ent8fX5/gIGCg4SFhoeIiYqLjI2Oj5CRkpOUlZaXmJmam5ydnp+goaKjpKWmp6ipqqusra6vsLGys7S1tre4ubq7vL2+v8DBwsPExcbHyMnKy8zNzs/Q0dLT1NXW19jZ2tvc3d7f4OHi4+Tl5ufo6err7O3u7/Dx8vM=") {
                throw "smallImage.value is not correct";
            }
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_stream_basic = test_stream_basic;
    function test_matrix_basic() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var testCaseObject = FakeExcelApi.TestCaseObject.newObject(ctx);
        var range1 = ctx.application.activeWorkbook.activeWorksheet.range("a1");
        range1.value = 100;
        var range2 = ctx.application.activeWorkbook.activeWorksheet.range("a2");
        range2.value = 200;
        var range3 = ctx.application.activeWorkbook.activeWorksheet.range("a3");
        range3.value = 300;
        var results = [];
        var expectedResults = [];
        var result = testCaseObject.sum(1, 2, 3);
        results.push(result);
        expectedResults.push(6);
        result = testCaseObject.sum(1, range1, range2);
        results.push(result);
        expectedResults.push(301);
        result = testCaseObject.sum(range3, 1, range2);
        results.push(result);
        expectedResults.push(300 + 1 + 200);
        result = testCaseObject.matrixSum([[1, 2, 3], [1, range1, range2], [range1, 2, 3]]);
        results.push(result);
        expectedResults.push(1 + 2 + 3 + 1 + 100 + 200 + 100 + 2 + 3);
        result = testCaseObject.matrixSum([[1, 2, range3], [1, range1, range2], [range1, 2, range3]]);
        results.push(result);
        expectedResults.push(1 + 2 + 300 + 1 + 100 + 200 + 100 + 2 + 300);
        result = testCaseObject.matrixSum([[range3, 2, range3], [1, range1, range2], [range1, [2, range2], range3]]);
        results.push(result);
        expectedResults.push(300 + 2 + 300 + 1 + 100 + 200 + 100 + 2 + +200 + 300);
        ctx.sync().then(function () {
            for (var i = 0; i < results.length; i++) {
                RichApiTest.log.comment("Test " + i + " Expected " + expectedResults[i] + ", Actual " + results[i].value);
                if (results[i].value != expectedResults[i]) {
                    throw "Test " + i + " Expected " + expectedResults[i] + ", Actual " + results[i].value;
                }
            }
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_matrix_basic = test_matrix_basic;
    function testReferencesAddRemoveArrayOfRefs() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var count0 = ctx.application.testWorkbook.getCachedObjectCount();
        ctx.sync().then(function () {
            RichApiTest.log.comment("StartCachedObjectCount = " + count0.value);
            var r1 = ctx.application.activeWorkbook.activeWorksheet.range("A1");
            var r2 = ctx.application.activeWorkbook.activeWorksheet.range("A2");
            var r3 = ctx.application.activeWorkbook.activeWorksheet.range("A3");
            var r4NotKept = ctx.application.activeWorkbook.activeWorksheet.range("A3");
            ctx.trackedObjects.add([r1, r2, r3]);
            return [r1, r2, r3];
        }).then(ctx.sync).then(function (refs) {
            return { refs: refs, count1: ctx.application.testWorkbook.getCachedObjectCount() };
        }).then(ctx.sync).then(function (passThrough) {
            RichApiTest.log.comment("CachedObjectCount with added references = " + passThrough.count1.value);
            // Three kept-references, plus for some reason cached object count counts everything as x2
            if ((passThrough.count1.value - count0.value) != 3 * 2) {
                throw new Error("Actual and expected counts do not match");
            }
            return passThrough.refs;
        }).then(function (refs) {
            ctx.trackedObjects.remove(refs);
            var finalCount = ctx.application.testWorkbook.getCachedObjectCount();
            return ctx.sync(finalCount);
        }).then(function (finalCount) {
            RichApiTest.log.comment("EndCachedObjectCount = " + finalCount.value);
            if (finalCount.value != count0.value) {
                throw new Error("Number of kept references should have dropped back down to its original value");
            }
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.testReferencesAddRemoveArrayOfRefs = testReferencesAddRemoveArrayOfRefs;
    function test_run_basic() {
        var run = FakeExcelApi.run(function (ctx) {
            var newText = "Hello, Keep Reference";
            var count0 = ctx.application.testWorkbook.getCachedObjectCount();
            return ctx.sync().then(function () {
                RichApiTest.log.comment("StartCachedObjectCount = " + count0.value);
                var r = ctx.application.activeWorkbook.activeWorksheet.range("A1");
                r.text = newText;
                return r;
            }).then(ctx.sync).then(function (r) {
                ctx.load(r);
                return { range: r, count1: ctx.application.testWorkbook.getCachedObjectCount() };
            }).then(ctx.sync).then(function (passThrough) {
                RichApiTest.log.comment("CachedObjectCount with added references = " + passThrough.count1.value);
                if (passThrough.count1.value <= count0.value) {
                    throw new Error("Expected count to be greater than initially, since should be keeping a reference on Range");
                }
                RichApiTest.log.comment("Range.text=" + passThrough.range.text);
                if (passThrough.range.text != newText) {
                    throw new Error("Text not equal to new text");
                }
            }).then(function () {
                // Return out the context, so can use it to validate the object count
                return { ctx: ctx, initialCount: count0.value };
            });
        });
        run.then(function (passedThrough) {
            RichApiTest.log.comment("Done with task");
            setTimeout(function () {
                RichApiTest.log.comment("Checking on the cleanup");
                // Let the clenaup fire (by waiting just a bit), and check on object count:
                var finalCount = passedThrough.ctx.application.testWorkbook.getCachedObjectCount();
                passedThrough.ctx.sync().then(function () {
                    RichApiTest.log.comment("Final CachedObjectCount = " + finalCount.value);
                    if (finalCount.value != passedThrough.initialCount) {
                        throw new Error("Final count (" + finalCount.value + ") does not match initial count (" + passedThrough.initialCount + ")");
                    }
                    RichApiTest.log.done(true);
                }).catch(FakeExcelTest.reportError);
            }, 1000);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_run_basic = test_run_basic;
    function test_run_previousAndSubsequentRefCountsGetReflected() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var count0 = ctx.application.testWorkbook.getCachedObjectCount();
        ctx.sync().then(function () {
            RichApiTest.log.comment("StartCachedObjectCount = " + count0.value);
        }).then(function () {
            ctx.trackedObjects.add(ctx.application.activeWorkbook.activeWorksheet.range("A1"));
        }).then(ctx.sync).then(function () { return FakeExcelApi.run(function (ctx) {
            var newText = "Hello, Keep Reference";
            var r = ctx.application.activeWorkbook.activeWorksheet.range("A1");
            r.text = newText;
            var count1 = ctx.application.testWorkbook.getCachedObjectCount();
            return ctx.sync().then(function () {
                RichApiTest.log.comment("CachedObjectCount with added references = " + count1.value);
                if (count1.value <= count0.value) {
                    throw new Error("Expected count to be greater than initially, since should be keeping a reference on Range");
                }
            });
        }); }).then(function () {
            RichApiTest.log.comment("Done with task");
            ctx.trackedObjects.add(ctx.application.activeWorkbook.activeWorksheet.range("A1"));
            ctx.trackedObjects.add(ctx.application.activeWorkbook.activeWorksheet.range("A2"));
            return ctx.sync().then(function () {
                setTimeout(function () {
                    RichApiTest.log.comment("Checking on the cleanup");
                    // Let the clenaup fire (by waiting just a bit), and check on object count:
                    var finalCount = ctx.application.testWorkbook.getCachedObjectCount();
                    ctx.sync().then(function () {
                        RichApiTest.log.comment("Final CachedObjectCount = " + finalCount.value);
                        var expectedDiffInCount = 6; // 1 for initial, 2 for final -- and both multiplied by 2, 
                        // since cachedObjectCount seems to reflect double the amount (maybe because it keeps alive a reference to parent worksheet?)
                        if (finalCount.value != (count0.value + expectedDiffInCount)) {
                            throw new Error("Final count (" + finalCount.value + ") does not match expected relative to initial count.");
                        }
                        RichApiTest.log.done(true);
                    }).catch(FakeExcelTest.reportError);
                }, 1000);
            });
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_run_previousAndSubsequentRefCountsGetReflected = test_run_previousAndSubsequentRefCountsGetReflected;
    function test_run_failsIfDontReturnPromise1() {
        FakeExcelApi.run(function (ctx) {
            var newText = "Hello, Keep Reference";
            var r = ctx.application.activeWorkbook.activeWorksheet.range("A1");
            r.text = newText;
            var count1 = ctx.application.testWorkbook.getCachedObjectCount();
            ctx.sync();
        }).then(function () {
            throw new Error("Should not have reached here");
        }).catch(function (e) {
            if ((e instanceof OfficeExtension.Error) && (e.code == "RunMustReturnPromise")) {
                RichApiTest.log.pass("Caught expected error");
            }
            else {
                FakeExcelTest.reportError(e);
            }
        });
    }
    FakeExcelTest.test_run_failsIfDontReturnPromise1 = test_run_failsIfDontReturnPromise1;
    function test_run_failsIfDontReturnPromise2() {
        FakeExcelApi.run(function (ctx) {
            var newText = "Hello, Keep Reference";
            var r = ctx.application.activeWorkbook.activeWorksheet.range("A1");
            r.text = newText;
            var count1 = ctx.application.testWorkbook.getCachedObjectCount();
            ctx.sync();
            return 5;
        }).then(function () {
            throw new Error("Should not have reached here");
        }).catch(function (e) {
            if ((e instanceof OfficeExtension.Error) && (e.code == "RunMustReturnPromise")) {
                RichApiTest.log.pass("Caught expected error");
            }
            else {
                FakeExcelTest.reportError(e);
            }
        });
    }
    FakeExcelTest.test_run_failsIfDontReturnPromise2 = test_run_failsIfDontReturnPromise2;
    function test_run_failsIfDontReturnPromise3() {
        FakeExcelApi.run(function (ctx) {
            var newText = "Hello, Keep Reference";
            var r = ctx.application.activeWorkbook.activeWorksheet.range("A1");
            r.text = newText;
            var count1 = ctx.application.testWorkbook.getCachedObjectCount();
            ctx.sync();
            return null;
        }).then(function () {
            throw new Error("Should not have reached here");
        }).catch(function (e) {
            if ((e instanceof OfficeExtension.Error) && (e.code == "RunMustReturnPromise")) {
                RichApiTest.log.pass("Caught expected error");
            }
            else {
                FakeExcelTest.reportError(e);
            }
        });
    }
    FakeExcelTest.test_run_failsIfDontReturnPromise3 = test_run_failsIfDontReturnPromise3;
    function test_run_throwsCorrectly1() {
        var EXPECTED_ERROR_TEXT = "Expected user-code error";
        FakeExcelApi.run(function (ctx) {
            throw new Error(EXPECTED_ERROR_TEXT);
            return ctx.sync();
        }).then(function () {
            throw new Error("Should not have reached here");
        }).catch(function (e) {
            if (e.message == EXPECTED_ERROR_TEXT) {
                RichApiTest.log.pass("Caught expected error");
            }
            else {
                FakeExcelTest.reportError(e);
            }
        });
    }
    FakeExcelTest.test_run_throwsCorrectly1 = test_run_throwsCorrectly1;
    function test_run_throwsCorrectly2() {
        var EXPECTED_ERROR_TEXT = "Expected user-code error";
        FakeExcelApi.run(function (ctx) {
            var r = ctx.application.activeWorkbook.activeWorksheet.range("A1");
            r.load("address");
            var count1 = ctx.application.testWorkbook.getCachedObjectCount();
            return ctx.sync().then(function () {
                throw new Error(EXPECTED_ERROR_TEXT);
            });
        }).then(function () {
            throw new Error("Should not have reached here");
        }).catch(function (e) {
            if (e.message == EXPECTED_ERROR_TEXT) {
                RichApiTest.log.pass("Caught expected error");
            }
            else {
                FakeExcelTest.reportError(e);
            }
        });
    }
    FakeExcelTest.test_run_throwsCorrectly2 = test_run_throwsCorrectly2;
    function test_run_invalidatesReferencedObject1() {
        var range;
        var run = FakeExcelApi.run(function (ctx) {
            var newText = "Hello, Keep Reference";
            range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
            range.text = newText;
            return ctx.sync();
        }).catch(function () {
            throw new Error("Should not have had any errors above");
        }).then(Util.wait(1000)).then(function () {
            range.load("address");
            return range.context.sync();
        }).catch(function (e) {
            if ((e instanceof OfficeExtension.Error) && (e.code == "InvalidObjectPath")) {
                RichApiTest.log.pass("Caught expected error");
            }
            else {
                FakeExcelTest.reportError(e);
            }
        });
    }
    FakeExcelTest.test_run_invalidatesReferencedObject1 = test_run_invalidatesReferencedObject1;
    function test_run_invalidatesReferencedObject2() {
        var range;
        var run = FakeExcelApi.run(function (ctx) {
            var newText = "Hello, Keep Reference";
            range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
            range.text = newText;
            // Create a timer to try to use the range, after it should have already gotten cleaned up
            setTimeout(function () {
                try {
                    range.load("address");
                    throw new Error("Should have exploded on the above call already");
                }
                catch (e) {
                    if ((e instanceof OfficeExtension.Error) && (e.code == "InvalidObjectPath")) {
                        RichApiTest.log.pass("Caught expected error");
                    }
                    else {
                        FakeExcelTest.reportError(e);
                    }
                }
            }, 1000);
            return ctx.sync();
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_run_invalidatesReferencedObject2 = test_run_invalidatesReferencedObject2;
    function test_run_invalidatesReferencedObject3() {
        // Same as test above, but using the internal API for onCleanupSuccess rather than waiting on timer job
        var range;
        var batch = function (ctx) {
            var newText = "Hello, Keep Reference";
            range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
            range.text = newText;
            return ctx.sync();
        };
        var onCleanupSuccess = function () {
            try {
                range.load("address");
                throw new Error("Should have exploded on the above call already");
            }
            catch (e) {
                if ((e instanceof OfficeExtension.Error) && (e.code == "InvalidObjectPath")) {
                    RichApiTest.log.pass("Caught expected error");
                }
                else {
                    FakeExcelTest.reportError(e);
                }
            }
        };
        var onCleanupFailure = function () {
            RichApiTest.log.fail("Cleanup failed, this should not have happened");
        };
        OfficeExtension.ClientRequestContext._run(function () { return new FakeExcelApi.ExcelClientRequestContext(); }, batch, 1, 5000, onCleanupSuccess, onCleanupFailure).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_run_invalidatesReferencedObject3 = test_run_invalidatesReferencedObject3;
    function test_run_canKeepReferenceExplicitly() {
        var range;
        var count0;
        var batch = function (ctx) {
            count0 = ctx.application.testWorkbook.getCachedObjectCount();
            var newText = "Hello, Keep Reference";
            range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
            range.text = newText;
            ctx.trackedObjects.add(range);
            return ctx.sync();
        };
        var onCleanupSuccess = function () {
            try {
                var ctx = range.context;
                RichApiTest.log.comment("Start CachedObjectCount = " + count0.value);
                range.load("address");
                var count1 = range.context.application.testWorkbook.getCachedObjectCount();
                var count2;
                ctx.trackedObjects.remove(range);
                ctx.sync().then(function () {
                    count2 = range.context.application.testWorkbook.getCachedObjectCount();
                }).then(ctx.sync).then(function () {
                    RichApiTest.log.comment("Intermediary CachedObjectCount = " + count1.value);
                    RichApiTest.log.comment("Final CachedObjectCount = " + count2.value);
                    if (count0.value != count2.value || count1.value <= count0.value) {
                        throw new Error("Initial reference count does not match final count, and/or didn't grow during the intermediary stage");
                    }
                    RichApiTest.log.pass("Was able to access the kept-referenced range object");
                }).catch(FakeExcelTest.reportError);
            }
            catch (e) {
                FakeExcelTest.reportError(e);
            }
        };
        var onCleanupFailure = function () {
            RichApiTest.log.fail("Cleanup failed, this should not have happened");
        };
        OfficeExtension.ClientRequestContext._run(function () { return new FakeExcelApi.ExcelClientRequestContext(); }, batch, 1, 10000, onCleanupSuccess, onCleanupFailure).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_run_canKeepReferenceExplicitly = test_run_canKeepReferenceExplicitly;
    function test_run_doesFinalSyncEvenIfUserForgot() {
        var range;
        FakeExcelApi.run(function (ctx) {
            var newText = "Hello, Keep Reference";
            range = ctx.application.activeWorkbook.activeWorksheet.range("A5");
            range.text = newText;
            return ctx.sync().then(function () {
                range.load("rowIndex");
                // Note that I ask to load, but specifically "forget" to call sync on it.
            });
        }).then(function () {
            RichApiTest.log.comment(range.rowIndex.toString());
        }).then(function () {
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_run_doesFinalSyncEvenIfUserForgot = test_run_doesFinalSyncEvenIfUserForgot;
    function test_run_emulateFailureOnInitialClenaups() {
        var cleanupCount = 0;
        var originalTime = performance.now();
        RichApiTest.log.comment("This test takes ~4 seconds, please be patient");
        var batch = function (ctx) {
            var newText = "Hello, Keep Reference";
            var range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
            range.text = newText;
            // Do some mucking with the cleanup routine:
            var originalRetrieveAndClear = ctx.trackedObjects._retrieveAndClearAutoCleanupList;
            ctx.trackedObjects._retrieveAndClearAutoCleanupList = function () {
                var originalSync = ctx.sync;
                ctx.sync = function (passedThroughValue) {
                    var isFirstOrSecondTry = (cleanupCount <= 1);
                    cleanupCount++;
                    if (isFirstOrSecondTry) {
                        return new OfficeExtension["Promise"](function (resolve, reject) {
                            reject(new Error("Emulating failing initially"));
                        });
                    }
                    else {
                        return originalSync(passedThroughValue);
                    }
                };
                return originalRetrieveAndClear();
            };
            return ctx.sync(ctx);
        };
        var onCleanupSuccess = function (attemptNumber) {
            if (attemptNumber <= 2) {
                RichApiTest.log.fail("Cleanup succeeded immediately, this should not have happened");
            }
            else if (attemptNumber === 3) {
                var timeElapsed = performance.now() - originalTime;
                RichApiTest.log.pass("Cleaned up on attempt #" + attemptNumber + ", after " + timeElapsed + " milliseconds");
            }
            else {
                RichApiTest.log.fail("Should have succeeded on 3rd attempt, so this should not have happened");
            }
        };
        var onCleanupFailure = function (attemptNumber) {
            if (attemptNumber <= 2) {
                RichApiTest.log.comment("Cleanup attempt #" + attemptNumber + " failed, as expected.  Wait for attempt #3 to pass.");
            }
            else {
                RichApiTest.log.fail("Cleanup didn't succeed on second try either, this should not have happened");
            }
        };
        OfficeExtension.ClientRequestContext._run(function () { return new FakeExcelApi.ExcelClientRequestContext(); }, batch, 3, 2000, onCleanupSuccess, onCleanupFailure).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_run_emulateFailureOnInitialClenaups = test_run_emulateFailureOnInitialClenaups;
    function test_run_emulateCleanupFailure() {
        var cleanupCount = 0;
        var originalTime = performance.now();
        RichApiTest.log.comment("This test takes ~8 seconds, please be patient");
        var batch = function (ctx) {
            var newText = "Hello, Keep Reference";
            var range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
            range.text = newText;
            // Do some mucking with the cleanup routine:
            var originalRetrieveAndClear = ctx.trackedObjects._retrieveAndClearAutoCleanupList;
            ctx.trackedObjects._retrieveAndClearAutoCleanupList = function () {
                var originalSync = ctx.sync;
                ctx.sync = function (passedThroughValue) {
                    cleanupCount++;
                    return new OfficeExtension["Promise"](function (resolve, reject) {
                        reject(new Error("Emulate failing"));
                    });
                };
                return originalRetrieveAndClear();
            };
            return ctx.sync(ctx);
        };
        var onCleanupSuccess = function (attemptNumber) {
            RichApiTest.log.fail("Should never succeed");
        };
        var onCleanupFailure = function (attemptNumber) {
            RichApiTest.log.comment("Cleanup attempt #" + attemptNumber + " failed, as expected. " + "Should try 2 times, then give up. Will validate a few seconds later, to make sure didn't keep on trying forever");
        };
        OfficeExtension.ClientRequestContext._run(function () { return new FakeExcelApi.ExcelClientRequestContext(); }, batch, 2, 2000, onCleanupSuccess, onCleanupFailure).catch(FakeExcelTest.reportError);
        setTimeout(function () {
            if (cleanupCount == 2) {
                RichApiTest.log.pass("Cleanup count as expected, tried twice then gave up");
            }
            else {
                RichApiTest.log.fail("Cleanup count doesn't match expected");
            }
        }, 8000);
    }
    FakeExcelTest.test_run_emulateCleanupFailure = test_run_emulateCleanupFailure;
    function testRestOnly_rest() {
        var request = {
            url: OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/activeWorksheet/RestOnlyOperation",
            method: RichApiTest.RestUtility.httpMethodGet
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            RichApiTest.log.done(true);
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    FakeExcelTest.testRestOnly_rest = testRestOnly_rest;
    function test_complextype_load_JScript() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        var sort = range.sort;
        ctx.load(sort);
        ctx.sync().then(function () {
            if (sort.queryField == null) {
                throw "queryField is null";
            }
            if (sort.queryField.rowLimit != 123) {
                throw "rowLimit != 123";
            }
            if (sort.queryField.field == null) {
                throw "queryField.field is null";
            }
            if (sort.queryField.field.columnIndex != 123) {
                throw "queryField.field.columnIndex != 123";
            }
            RichApiTest.log.comment(JSON.stringify(sort.queryField));
            if (sort.fields == null) {
                throw "fields is null";
            }
            if (sort.fields.length != 5) {
                throw "fields.length != 5";
            }
            RichApiTest.log.comment(JSON.stringify(sort.fields));
        }).then(function () {
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_complextype_load_JScript = test_complextype_load_JScript;
    function test_complextype_get_REST() {
        var url = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/activeWorksheet/range('A1')/sort";
        var request = {
            url: url,
            method: RichApiTest.RestUtility.httpMethodGet,
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            var sort = JSON.parse(resp.body);
            if (!sort.queryField) {
                throw "!sort.queryField";
            }
            if (!sort.fields) {
                throw "!sort.fields";
            }
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_complextype_get_REST = test_complextype_get_REST;
    function test_complextype_arrayinput_JScript() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        var sort = range.sort;
        var result = sort.apply([{ columnIndex: 100, assending: true }, { columnIndex: 101, assending: false }]);
        ctx.sync().then(function () {
            RichApiTest.log.comment(JSON.stringify(result.value));
            if (result.value != "100,101,") {
                throw "result.value != 100,101";
            }
        }).then(function () {
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_complextype_arrayinput_JScript = test_complextype_arrayinput_JScript;
    function test_complextype_arrayinput_REST() {
        var input = { "fields": [{ columnIndex: 100, assending: true }, { columnIndex: 101, assending: false }] };
        var url = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/activeWorksheet/range('A1')/sort/apply";
        var request = {
            url: url,
            method: RichApiTest.RestUtility.httpMethodPost,
            body: JSON.stringify(input)
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            var result = JSON.parse(resp.body);
            if (result.value != "100,101,") {
                throw "result.value != 100,101";
            }
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_complextype_arrayinput_REST = test_complextype_arrayinput_REST;
    function test_complextype_mixedArrayInput_JScript() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        var sort = range.sort;
        var result = sort.applyMixed([1, 2, { columnIndex: 100, assending: true }, 3, { columnIndex: 101, assending: false }]);
        ctx.sync().then(function () {
            RichApiTest.log.comment(JSON.stringify(result.value));
            if (result.value != 1 + 2 + 3 + 100 + 101) {
                throw "result.value != 207";
            }
        }).then(function () {
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_complextype_mixedArrayInput_JScript = test_complextype_mixedArrayInput_JScript;
    function test_complextype_mixedArrayInputODataType_JScript() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        var sort = range.sort;
        var result = sort.applyMixed([1, 2, { '@odata.type': 'ExcelApi.SortField', columnIndex: 100, assending: true }, 3, { '@odata.type': 'ExcelApi.SortField', columnIndex: 101, assending: false }]);
        ctx.sync().then(function () {
            RichApiTest.log.comment(JSON.stringify(result.value));
            if (result.value != 1 + 2 + 3 + 100 + 101) {
                throw "result.value != 207";
            }
        }).then(function () {
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_complextype_mixedArrayInputODataType_JScript = test_complextype_mixedArrayInputODataType_JScript;
    function test_complextype_mixedArrayInput_REST() {
        var input = { "fields": [1, 2, { columnIndex: 100, assending: true }, 3, { columnIndex: 101, assending: false }] };
        var url = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/activeWorksheet/range('A1')/sort/applyMixed";
        var request = {
            url: url,
            method: RichApiTest.RestUtility.httpMethodPost,
            body: JSON.stringify(input)
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            var result = JSON.parse(resp.body);
            if (result.value != 1 + 2 + 3 + 100 + 101) {
                throw "result.value != 207";
            }
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_complextype_mixedArrayInput_REST = test_complextype_mixedArrayInput_REST;
    function test_complextype_propertyset_JScript() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        var sort = range.sort;
        sort.fields2 = [{ columnIndex: 100, assending: true }, { columnIndex: 101, assending: false }];
        ctx.sync().then(function () {
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_complextype_propertyset_JScript = test_complextype_propertyset_JScript;
    function test_complextype_methodreturn_JScript() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        var sort = range.sort;
        var result = sort.applyAndReturnFirstField([{ columnIndex: 100, assending: true }, { columnIndex: 101, assending: false }]);
        ctx.sync().then(function () {
            RichApiTest.log.comment(JSON.stringify(result.value));
            if (result.value.columnIndex != 100) {
                throw "columnIndex != 100";
            }
        }).then(function () {
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_complextype_methodreturn_JScript = test_complextype_methodreturn_JScript;
    function test_complextype_methodreturn_REST() {
        var input = { "fields": [{ columnIndex: 100, assending: true }, { columnIndex: 101, assending: false }] };
        var url = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/activeWorksheet/range('A1')/sort/applyAndReturnFirstField";
        var request = {
            url: url,
            method: RichApiTest.RestUtility.httpMethodPost,
            body: JSON.stringify(input)
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            var result = JSON.parse(resp.body);
            if (result.value.columnIndex != 100) {
                throw "result.value.columnIndex != 100";
            }
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_complextype_methodreturn_REST = test_complextype_methodreturn_REST;
    function test_complextype_methodreturn2_JScript() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        var range = ctx.application.activeWorkbook.activeWorksheet.range("A1");
        var sort = range.sort;
        var result = sort.applyQueryWithSortFieldAndReturnLast([
            { rowLimit: 200, field: { columnIndex: 201, assending: true } },
            { rowLimit: 300, field: { columnIndex: 301, assending: false } }
        ]);
        ctx.sync().then(function () {
            RichApiTest.log.comment(JSON.stringify(result.value));
            if (result.value.field.columnIndex != 301) {
                throw "columnIndex != 301";
            }
        }).then(function () {
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_complextype_methodreturn2_JScript = test_complextype_methodreturn2_JScript;
    function test_complextype_methodreturn2_REST() {
        var input = {
            "fields": [
                { rowLimit: 200, field: { columnIndex: 201, assending: true } },
                { rowLimit: 300, field: { columnIndex: 301, assending: false } }
            ]
        };
        var url = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/activeWorksheet/range('A1')/sort/applyQueryWithSortFieldAndReturnLast";
        var request = {
            url: url,
            method: RichApiTest.RestUtility.httpMethodPost,
            body: JSON.stringify(input)
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            var result = JSON.parse(resp.body);
            if (result.value.field.columnIndex != 301) {
                throw "result.value.field.columnIndex != 301";
            }
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_complextype_methodreturn2_REST = test_complextype_methodreturn2_REST;
    function test_range_parameter_REST() {
        var url = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/sheets/findsheet(text='abc')/range('A1')";
        var request = {
            url: url,
            method: RichApiTest.RestUtility.httpMethodGet
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_range_parameter_REST = test_range_parameter_REST;
    function test_range_parameter2_REST() {
        var url = OfficeExtension.Constants.localDocumentApiPrefix + "activeWorkbook/$/sheets/$/findsheet(text='abc')/$/range('A1')";
        var request = {
            url: url,
            method: RichApiTest.RestUtility.httpMethodGet
        };
        RichApiTest.RestUtility.invoke(request).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_range_parameter2_REST = test_range_parameter2_REST;
    function test_application_hasbase_JScript() {
        var ctx = new FakeExcelApi.ExcelClientRequestContext();
        ctx.application.load("hasBase");
        ctx.sync().then(function () {
            if (!ctx.application.hasBase) {
                throw new Error("Application.hasBase is false");
            }
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_application_hasbase_JScript = test_application_hasbase_JScript;
    function test_Promise_exposedPromiseWorks_JScript() {
        new OfficeExtension.Promise(function (resolve, reject) {
            setTimeout(function () {
                resolve("500 waited");
            }, 500);
        }).then(function (value) {
            if (value != "500 waited") {
                throw new Error("Did not resolve correctly");
            }
            RichApiTest.log.done(true);
        }).catch(FakeExcelTest.reportError);
    }
    FakeExcelTest.test_Promise_exposedPromiseWorks_JScript = test_Promise_exposedPromiseWorks_JScript;
    var Util;
    (function (Util) {
        function wait(milliseconds) {
            return function () { return new OfficeExtension.Promise(function (resolve, reject) {
                setTimeout(function () {
                    resolve();
                }, milliseconds);
            }); };
        }
        Util.wait = wait;
    })(Util || (Util = {}));
})(FakeExcelTest || (FakeExcelTest = {}));
var OfficeExtension;
(function (OfficeExtension) {
    var FakeResponseRequestExecutor = (function () {
        function FakeResponseRequestExecutor(responseText) {
            this.m_responseText = responseText;
        }
        FakeResponseRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage, callback) {
            var response = { ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
            response.Body = JSON.parse(this.m_responseText);
            window.setTimeout(function () {
                callback(response);
            }, 100);
        };
        return FakeResponseRequestExecutor;
    })();
    OfficeExtension.FakeResponseRequestExecutor = FakeResponseRequestExecutor;
})(OfficeExtension || (OfficeExtension = {}));
/*
 * This is a generated file.  Generated by osfclient\RichApi\Test\FakeXlapiMetadata\FakeXlapiGen.bat.
 * If there are content placeholders, only edit content inside content placeholders.
 * If there are no content placeholders, do not edit this file directly.
 */
/* Begin_PlaceHolder_GlobalHeader */
/* End_PlaceHolder_GlobalHeader */
var __extends = this.__extends || function (d, b) {
    for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p];
    function __() { this.constructor = d; }
    __.prototype = b.prototype;
    d.prototype = new __();
};
var FakeExcelApi;
(function (FakeExcelApi) {
    /* Begin_PlaceHolder_ModuleHeader */
    /* End_PlaceHolder_ModuleHeader */
    var _createPropertyObjectPath = OfficeExtension.ObjectPathFactory.createPropertyObjectPath;
    var _createMethodObjectPath = OfficeExtension.ObjectPathFactory.createMethodObjectPath;
    var _createIndexerObjectPath = OfficeExtension.ObjectPathFactory.createIndexerObjectPath;
    var _createNewObjectObjectPath = OfficeExtension.ObjectPathFactory.createNewObjectObjectPath;
    var _createChildItemObjectPathUsingIndexer = OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexer;
    var _createChildItemObjectPathUsingGetItemAt = OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingGetItemAt;
    var _createChildItemObjectPathUsingIndexerOrGetItemAt = OfficeExtension.ObjectPathFactory.createChildItemObjectPathUsingIndexerOrGetItemAt;
    var _createMethodAction = OfficeExtension.ActionFactory.createMethodAction;
    var _createSetPropertyAction = OfficeExtension.ActionFactory.createSetPropertyAction;
    var _isNullOrUndefined = OfficeExtension.Utility.isNullOrUndefined;
    var _isUndefined = OfficeExtension.Utility.isUndefined;
    var _throwIfNotLoaded = OfficeExtension.Utility.throwIfNotLoaded;
    var _load = OfficeExtension.Utility.load;
    var _fixObjectPathIfNecessary = OfficeExtension.Utility.fixObjectPathIfNecessary;
    var _addActionResultHandler = OfficeExtension.Utility._addActionResultHandler;
    var _handleNavigationPropertyResults = OfficeExtension.Utility._handleNavigationPropertyResults;
    /**
     *
     * Chart type
     */
    var ChartType;
    (function (ChartType) {
        /**
         *
         * Not specified
         */
        ChartType.none = "None";
        /**
         *
         * Pie chart
         */
        ChartType.pie = "Pie";
        /**
         *
         * Bar chart
         */
        ChartType.bar = "Bar";
        /**
         *
         * Line chart
         */
        ChartType.line = "Line";
        /**
         *
         * 3D chart
         */
        ChartType._3DBar = "3DBar";
    })(ChartType = FakeExcelApi.ChartType || (FakeExcelApi.ChartType = {}));
    var RangeValueType;
    (function (RangeValueType) {
        RangeValueType.unknown = "Unknown";
        RangeValueType.empty = "Empty";
        RangeValueType.string = "String";
        RangeValueType.integer = "Integer";
        RangeValueType.double = "Double";
        RangeValueType.boolean = "Boolean";
        RangeValueType.error = "Error";
    })(RangeValueType = FakeExcelApi.RangeValueType || (FakeExcelApi.RangeValueType = {}));
    var Application = (function (_super) {
        __extends(Application, _super);
        function Application() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Application.prototype, "activeWorkbook", {
            /* Begin_PlaceHolder_Application_Custom_Members */
            /* End_PlaceHolder_Application_Custom_Members */
            get: function () {
                if (!this.m_activeWorkbook) {
                    this.m_activeWorkbook = new FakeExcelApi.Workbook(this.context, _createPropertyObjectPath(this.context, this, "ActiveWorkbook", false, false));
                }
                return this.m_activeWorkbook;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "testWorkbook", {
            get: function () {
                if (!this.m_testWorkbook) {
                    this.m_testWorkbook = new FakeExcelApi.TestWorkbook(this.context, _createPropertyObjectPath(this.context, this, "TestWorkbook", false, false));
                }
                return this.m_testWorkbook;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Application.prototype, "hasBase", {
            get: function () {
                _throwIfNotLoaded("hasBase", this.m_hasBase);
                return this.m_hasBase;
            },
            enumerable: true,
            configurable: true
        });
        Application.prototype._GetObjectByReferenceId = function (referenceId) {
            /* Begin_PlaceHolder_Application__GetObjectByReferenceId */
            /* End_PlaceHolder_Application__GetObjectByReferenceId */
            var action = _createMethodAction(this.context, this, "_GetObjectByReferenceId", 1 /* Read */, [referenceId]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        Application.prototype._GetObjectTypeNameByReferenceId = function (referenceId) {
            /* Begin_PlaceHolder_Application__GetObjectTypeNameByReferenceId */
            /* End_PlaceHolder_Application__GetObjectTypeNameByReferenceId */
            var action = _createMethodAction(this.context, this, "_GetObjectTypeNameByReferenceId", 1 /* Read */, [referenceId]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        Application.prototype._RemoveReference = function (referenceId) {
            /* Begin_PlaceHolder_Application__RemoveReference */
            /* End_PlaceHolder_Application__RemoveReference */
            _createMethodAction(this.context, this, "_RemoveReference", 1 /* Read */, [referenceId]);
        };
        /** Handle results returned from the document
         * @private
         */
        Application.prototype._handleResult = function (value) {
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["HasBase"])) {
                this.m_hasBase = obj["HasBase"];
            }
            _handleNavigationPropertyResults(this, obj, ["activeWorkbook", "ActiveWorkbook", "testWorkbook", "TestWorkbook"]);
        };
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        Application.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return Application;
    })(OfficeExtension.ClientObject);
    FakeExcelApi.Application = Application;
    var Chart = (function (_super) {
        __extends(Chart, _super);
        function Chart() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Chart.prototype, "chartType", {
            /* Begin_PlaceHolder_Chart_Custom_Members */
            /* End_PlaceHolder_Chart_Custom_Members */
            get: function () {
                _throwIfNotLoaded("chartType", this.m_chartType);
                return this.m_chartType;
            },
            set: function (value) {
                this.m_chartType = value;
                _createSetPropertyAction(this.context, this, "ChartType", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "id", {
            get: function () {
                _throwIfNotLoaded("id", this.m_id);
                return this.m_id;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "imageData", {
            get: function () {
                _throwIfNotLoaded("imageData", this.m_imageData);
                return this.m_imageData;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name);
                return this.m_name;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "nullableChartType", {
            get: function () {
                _throwIfNotLoaded("nullableChartType", this.m_nullableChartType);
                return this.m_nullableChartType;
            },
            set: function (value) {
                this.m_nullableChartType = value;
                _createSetPropertyAction(this.context, this, "NullableChartType", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "nullableShowLabel", {
            get: function () {
                _throwIfNotLoaded("nullableShowLabel", this.m_nullableShowLabel);
                return this.m_nullableShowLabel;
            },
            set: function (value) {
                this.m_nullableShowLabel = value;
                _createSetPropertyAction(this.context, this, "NullableShowLabel", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Chart.prototype, "title", {
            get: function () {
                _throwIfNotLoaded("title", this.m_title);
                return this.m_title;
            },
            set: function (value) {
                this.m_title = value;
                _createSetPropertyAction(this.context, this, "Title", value);
            },
            enumerable: true,
            configurable: true
        });
        /**
         *
         * Delete the chart
         *
         */
        Chart.prototype.delete = function () {
            /* Begin_PlaceHolder_Chart_Delete */
            /* End_PlaceHolder_Chart_Delete */
            _createMethodAction(this.context, this, "Delete", 0 /* Default */, []);
        };
        Chart.prototype.getAsImage = function (large) {
            /* Begin_PlaceHolder_Chart_GetAsImage */
            /* End_PlaceHolder_Chart_GetAsImage */
            var action = _createMethodAction(this.context, this, "GetAsImage", 1 /* Read */, [large]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        /** Handle results returned from the document
         * @private
         */
        Chart.prototype._handleResult = function (value) {
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["ChartType"])) {
                this.m_chartType = obj["ChartType"];
            }
            if (!_isUndefined(obj["Id"])) {
                this.m_id = obj["Id"];
            }
            if (!_isUndefined(obj["ImageData"])) {
                this.m_imageData = obj["ImageData"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["NullableChartType"])) {
                this.m_nullableChartType = obj["NullableChartType"];
            }
            if (!_isUndefined(obj["NullableShowLabel"])) {
                this.m_nullableShowLabel = obj["NullableShowLabel"];
            }
            if (!_isUndefined(obj["Title"])) {
                this.m_title = obj["Title"];
            }
        };
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        Chart.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return Chart;
    })(OfficeExtension.ClientObject);
    FakeExcelApi.Chart = Chart;
    /**
     *
     * Chart collection
     */
    var ChartCollection = (function (_super) {
        __extends(ChartCollection, _super);
        function ChartCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(ChartCollection.prototype, "items", {
            /* Begin_PlaceHolder_ChartCollection_Custom_Members */
            /* End_PlaceHolder_ChartCollection_Custom_Members */
            /** Gets the loaded child items in this collection. */
            get: function () {
                _throwIfNotLoaded("items", this.m__items);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(ChartCollection.prototype, "count", {
            /**
             *
             * Gets the number of charts
             */
            get: function () {
                _throwIfNotLoaded("count", this.m_count);
                return this.m_count;
            },
            enumerable: true,
            configurable: true
        });
        /**
         *
         * Add a new chart to the collection
         *
         * @param name The name of the chart to be added
         * @param chartType The type of the chart to be added
         * @returns The newly added chart
         */
        ChartCollection.prototype.add = function (name, chartType) {
            /* Begin_PlaceHolder_ChartCollection_Add */
            /* End_PlaceHolder_ChartCollection_Add */
            return new FakeExcelApi.Chart(this.context, _createMethodObjectPath(this.context, this, "Add", 0 /* Default */, [name, chartType], false, true));
        };
        /**
         *
         * Get the chart at the index
         *
         * @param index The index
         * @returns The chart at the index
         */
        ChartCollection.prototype.getItem = function (index) {
            /* Begin_PlaceHolder_ChartCollection_GetItem */
            /* End_PlaceHolder_ChartCollection_GetItem */
            return new FakeExcelApi.Chart(this.context, _createIndexerObjectPath(this.context, this, [index]));
        };
        /**
         *
         * Get the chart at the position
         *
         * @param ordinal The position
         * @returns The chart at the position
         */
        ChartCollection.prototype.getItemAt = function (ordinal) {
            /* Begin_PlaceHolder_ChartCollection_GetItemAt */
            /* End_PlaceHolder_ChartCollection_GetItemAt */
            return new FakeExcelApi.Chart(this.context, _createMethodObjectPath(this.context, this, "GetItemAt", 1 /* Read */, [ordinal], false, false));
        };
        /** Handle results returned from the document
         * @private
         */
        ChartCollection.prototype._handleResult = function (value) {
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Count"])) {
                this.m_count = obj["Count"];
            }
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new FakeExcelApi.Chart(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        ChartCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return ChartCollection;
    })(OfficeExtension.ClientObject);
    FakeExcelApi.ChartCollection = ChartCollection;
    var ErrorMethodType;
    (function (ErrorMethodType) {
        ErrorMethodType.none = "None";
        ErrorMethodType.accessDenied = "AccessDenied";
        ErrorMethodType.stateChanged = "StateChanged";
        ErrorMethodType.bounds = "Bounds";
        ErrorMethodType.abort = "Abort";
    })(ErrorMethodType = FakeExcelApi.ErrorMethodType || (FakeExcelApi.ErrorMethodType = {}));
    var TestWorkbook = (function (_super) {
        __extends(TestWorkbook, _super);
        function TestWorkbook() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(TestWorkbook.prototype, "errorWorksheet", {
            /* Begin_PlaceHolder_TestWorkbook_Custom_Members */
            /* End_PlaceHolder_TestWorkbook_Custom_Members */
            /**
             *
             * When this property is accessed, the server will return errorCode E_CHANGED_STATE
             */
            get: function () {
                if (!this.m_errorWorksheet) {
                    this.m_errorWorksheet = new FakeExcelApi.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "ErrorWorksheet", false, false));
                }
                return this.m_errorWorksheet;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(TestWorkbook.prototype, "errorWorksheet2", {
            /**
             *
             * When this property is accessed, the server will return errorCode
             */
            get: function () {
                if (!this.m_errorWorksheet2) {
                    this.m_errorWorksheet2 = new FakeExcelApi.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "ErrorWorksheet2", false, false));
                }
                return this.m_errorWorksheet2;
            },
            enumerable: true,
            configurable: true
        });
        /**
         *
         * When this method is invoked, the server will return errorCode. The errorCode is dependent on the ErrorMethodType input.
         *
         * @param input
         * @returns
         */
        TestWorkbook.prototype.errorMethod = function (input) {
            /* Begin_PlaceHolder_TestWorkbook_ErrorMethod */
            /* End_PlaceHolder_TestWorkbook_ErrorMethod */
            var action = _createMethodAction(this.context, this, "ErrorMethod", 0 /* Default */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        /**
         *
         * When this method is invoked, the server will return errorCode. The errorCode is dependent on the ErrorMethodType input.
         *
         * @param input
         * @returns
         */
        TestWorkbook.prototype.errorMethod2 = function (input) {
            /* Begin_PlaceHolder_TestWorkbook_ErrorMethod2 */
            /* End_PlaceHolder_TestWorkbook_ErrorMethod2 */
            var action = _createMethodAction(this.context, this, "ErrorMethod2", 0 /* Default */, [input]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        TestWorkbook.prototype.getActiveWorksheet = function () {
            /* Begin_PlaceHolder_TestWorkbook_GetActiveWorksheet */
            /* End_PlaceHolder_TestWorkbook_GetActiveWorksheet */
            return new FakeExcelApi.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetActiveWorksheet", 1 /* Read */, [], false, false));
        };
        /**
         *
         * Get the total count of active cached range data objects
         *
         * @returns
         */
        TestWorkbook.prototype.getCachedObjectCount = function () {
            /* Begin_PlaceHolder_TestWorkbook_GetCachedObjectCount */
            /* End_PlaceHolder_TestWorkbook_GetCachedObjectCount */
            var action = _createMethodAction(this.context, this, "GetCachedObjectCount", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        TestWorkbook.prototype.getNullableBoolValue = function (nullable) {
            /* Begin_PlaceHolder_TestWorkbook_GetNullableBoolValue */
            /* End_PlaceHolder_TestWorkbook_GetNullableBoolValue */
            var action = _createMethodAction(this.context, this, "GetNullableBoolValue", 0 /* Default */, [nullable]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        TestWorkbook.prototype.getNullableEnumValue = function (nullable) {
            /* Begin_PlaceHolder_TestWorkbook_GetNullableEnumValue */
            /* End_PlaceHolder_TestWorkbook_GetNullableEnumValue */
            var action = _createMethodAction(this.context, this, "GetNullableEnumValue", 0 /* Default */, [nullable]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        /**
         *
         * Get the total count of active COM objects
         *
         * @returns
         */
        TestWorkbook.prototype.getObjectCount = function () {
            /* Begin_PlaceHolder_TestWorkbook_GetObjectCount */
            /* End_PlaceHolder_TestWorkbook_GetObjectCount */
            var action = _createMethodAction(this.context, this, "GetObjectCount", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        TestWorkbook.prototype.testNullableInputValue = function (chartType, boolValue) {
            /* Begin_PlaceHolder_TestWorkbook_TestNullableInputValue */
            /* End_PlaceHolder_TestWorkbook_TestNullableInputValue */
            var action = _createMethodAction(this.context, this, "TestNullableInputValue", 0 /* Default */, [chartType, boolValue]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        /** Handle results returned from the document
         * @private
         */
        TestWorkbook.prototype._handleResult = function (value) {
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["errorWorksheet", "ErrorWorksheet", "errorWorksheet2", "ErrorWorksheet2"]);
        };
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        TestWorkbook.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return TestWorkbook;
    })(OfficeExtension.ClientObject);
    FakeExcelApi.TestWorkbook = TestWorkbook;
    var Workbook = (function (_super) {
        __extends(Workbook, _super);
        function Workbook() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(Workbook.prototype, "activeWorksheet", {
            /* Begin_PlaceHolder_Workbook_Custom_Members */
            /* End_PlaceHolder_Workbook_Custom_Members */
            get: function () {
                if (!this.m_activeWorksheet) {
                    this.m_activeWorksheet = new FakeExcelApi.Worksheet(this.context, _createPropertyObjectPath(this.context, this, "ActiveWorksheet", false, false));
                }
                return this.m_activeWorksheet;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "charts", {
            get: function () {
                if (!this.m_charts) {
                    this.m_charts = new FakeExcelApi.ChartCollection(this.context, _createPropertyObjectPath(this.context, this, "Charts", true, false));
                }
                return this.m_charts;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Workbook.prototype, "sheets", {
            get: function () {
                if (!this.m_sheets) {
                    this.m_sheets = new FakeExcelApi.WorksheetCollection(this.context, _createPropertyObjectPath(this.context, this, "Sheets", true, false));
                }
                return this.m_sheets;
            },
            enumerable: true,
            configurable: true
        });
        Workbook.prototype.getChartByType = function (chartType) {
            /* Begin_PlaceHolder_Workbook_GetChartByType */
            /* End_PlaceHolder_Workbook_GetChartByType */
            return new FakeExcelApi.Chart(this.context, _createMethodObjectPath(this.context, this, "GetChartByType", 1 /* Read */, [chartType], false, false));
        };
        Workbook.prototype.getChartByTypeTitle = function (chartType, title) {
            /* Begin_PlaceHolder_Workbook_GetChartByTypeTitle */
            /* End_PlaceHolder_Workbook_GetChartByTypeTitle */
            return new FakeExcelApi.Chart(this.context, _createMethodObjectPath(this.context, this, "GetChartByTypeTitle", 1 /* Read */, [chartType, title], false, false));
        };
        Workbook.prototype.someAction = function (intVal, strVal, enumVal) {
            /* Begin_PlaceHolder_Workbook_SomeAction */
            /* End_PlaceHolder_Workbook_SomeAction */
            var action = _createMethodAction(this.context, this, "SomeAction", 0 /* Default */, [intVal, strVal, enumVal]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        /** Handle results returned from the document
         * @private
         */
        Workbook.prototype._handleResult = function (value) {
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            _handleNavigationPropertyResults(this, obj, ["activeWorksheet", "ActiveWorksheet", "charts", "Charts", "sheets", "Sheets"]);
        };
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        Workbook.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return Workbook;
    })(OfficeExtension.ClientObject);
    FakeExcelApi.Workbook = Workbook;
    var Worksheet = (function (_super) {
        __extends(Worksheet, _super);
        function Worksheet() {
            _super.apply(this, arguments);
        }
        /* Begin_PlaceHolder_Worksheet_Custom_Members */
        /* End_PlaceHolder_Worksheet_Custom_Members */
        Worksheet.prototype.getActiveCell = function () {
            return new FakeExcelApi.Range(this.context, _createPropertyObjectPath(this.context, this, "ActiveCell", false, false));
        };
        Worksheet.prototype.getActiveCellInvalidAfterRequest = function () {
            return new FakeExcelApi.Range(this.context, _createPropertyObjectPath(this.context, this, "ActiveCellInvalidAfterRequest", false, true));
        };
        Object.defineProperty(Worksheet.prototype, "calculatedName", {
            get: function () {
                _throwIfNotLoaded("calculatedName", this.m_calculatedName);
                return this.m_calculatedName;
            },
            set: function (value) {
                this.m_calculatedName = value;
                _createSetPropertyAction(this.context, this, "CalculatedName", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "name", {
            get: function () {
                _throwIfNotLoaded("name", this.m_name);
                return this.m_name;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Worksheet.prototype, "_Id", {
            get: function () {
                _throwIfNotLoaded("_Id", this.m__Id);
                return this.m__Id;
            },
            enumerable: true,
            configurable: true
        });
        Worksheet.prototype.nullChart = function (address) {
            /* Begin_PlaceHolder_Worksheet_NullChart */
            /* End_PlaceHolder_Worksheet_NullChart */
            return new FakeExcelApi.Chart(this.context, _createMethodObjectPath(this.context, this, "NullChart", 1 /* Read */, [address], false, false));
        };
        Worksheet.prototype.nullRange = function (address) {
            /* Begin_PlaceHolder_Worksheet_NullRange */
            /* End_PlaceHolder_Worksheet_NullRange */
            return new FakeExcelApi.Range(this.context, _createMethodObjectPath(this.context, this, "NullRange", 1 /* Read */, [address], false, false));
        };
        Worksheet.prototype.range = function (address) {
            /* Begin_PlaceHolder_Worksheet_Range */
            /* End_PlaceHolder_Worksheet_Range */
            return new FakeExcelApi.Range(this.context, _createMethodObjectPath(this.context, this, "Range", 1 /* Read */, [address], false, true));
        };
        Worksheet.prototype.someRangeOperation = function (input, range) {
            /* Begin_PlaceHolder_Worksheet_SomeRangeOperation */
            /* End_PlaceHolder_Worksheet_SomeRangeOperation */
            var action = _createMethodAction(this.context, this, "SomeRangeOperation", 0 /* Default */, [input, range]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        Worksheet.prototype._RestOnly = function () {
            /* Begin_PlaceHolder_Worksheet__RestOnly */
            /* End_PlaceHolder_Worksheet__RestOnly */
            var action = _createMethodAction(this.context, this, "_RestOnly", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        /** Handle results returned from the document
         * @private
         */
        Worksheet.prototype._handleResult = function (value) {
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["CalculatedName"])) {
                this.m_calculatedName = obj["CalculatedName"];
            }
            if (!_isUndefined(obj["Name"])) {
                this.m_name = obj["Name"];
            }
            if (!_isUndefined(obj["_Id"])) {
                this.m__Id = obj["_Id"];
            }
        };
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        Worksheet.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return Worksheet;
    })(OfficeExtension.ClientObject);
    FakeExcelApi.Worksheet = Worksheet;
    var WorksheetCollection = (function (_super) {
        __extends(WorksheetCollection, _super);
        function WorksheetCollection() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(WorksheetCollection.prototype, "items", {
            /* Begin_PlaceHolder_WorksheetCollection_Custom_Members */
            /* End_PlaceHolder_WorksheetCollection_Custom_Members */
            /** Gets the loaded child items in this collection. */
            get: function () {
                _throwIfNotLoaded("items", this.m__items);
                return this.m__items;
            },
            enumerable: true,
            configurable: true
        });
        WorksheetCollection.prototype.add = function (name) {
            /* Begin_PlaceHolder_WorksheetCollection_Add */
            /* End_PlaceHolder_WorksheetCollection_Add */
            return new FakeExcelApi.Worksheet(this.context, _createMethodObjectPath(this.context, this, "Add", 0 /* Default */, [name], false, true));
        };
        WorksheetCollection.prototype.findSheet = function (text) {
            /* Begin_PlaceHolder_WorksheetCollection_FindSheet */
            /* End_PlaceHolder_WorksheetCollection_FindSheet */
            return new FakeExcelApi.Worksheet(this.context, _createMethodObjectPath(this.context, this, "FindSheet", 1 /* Read */, [text], false, false));
        };
        WorksheetCollection.prototype.getActiveWorksheetInvalidAfterRequest = function () {
            /* Begin_PlaceHolder_WorksheetCollection_GetActiveWorksheetInvalidAfterRequest */
            /* End_PlaceHolder_WorksheetCollection_GetActiveWorksheetInvalidAfterRequest */
            return new FakeExcelApi.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetActiveWorksheetInvalidAfterRequest", 1 /* Read */, [], false, true));
        };
        WorksheetCollection.prototype.getItem = function (index) {
            /* Begin_PlaceHolder_WorksheetCollection_GetItem */
            /* End_PlaceHolder_WorksheetCollection_GetItem */
            return new FakeExcelApi.Worksheet(this.context, _createMethodObjectPath(this.context, this, "GetItem", 0 /* Default */, [index], false, false));
        };
        WorksheetCollection.prototype._GetItem = function (index) {
            /* Begin_PlaceHolder_WorksheetCollection__GetItem */
            /* End_PlaceHolder_WorksheetCollection__GetItem */
            return new FakeExcelApi.Worksheet(this.context, _createIndexerObjectPath(this.context, this, [index]));
        };
        /** Handle results returned from the document
         * @private
         */
        WorksheetCollection.prototype._handleResult = function (value) {
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isNullOrUndefined(obj[OfficeExtension.Constants.items])) {
                this.m__items = [];
                var _data = obj[OfficeExtension.Constants.items];
                for (var i = 0; i < _data.length; i++) {
                    var _item = new FakeExcelApi.Worksheet(this.context, _createChildItemObjectPathUsingIndexerOrGetItemAt(true, this.context, this, _data[i], i));
                    _item._handleResult(_data[i]);
                    this.m__items.push(_item);
                }
            }
        };
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        WorksheetCollection.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return WorksheetCollection;
    })(OfficeExtension.ClientObject);
    FakeExcelApi.WorksheetCollection = WorksheetCollection;
    var Range = (function (_super) {
        __extends(Range, _super);
        function Range() {
            _super.apply(this, arguments);
        }
        /* Begin_PlaceHolder_Range_Custom_Members */
        Range.prototype.someCustomMethod = function () {
            console.log("someCustomMethod");
        };
        Object.defineProperty(Range.prototype, "sort", {
            /* End_PlaceHolder_Range_Custom_Members */
            get: function () {
                if (!this.m_sort) {
                    this.m_sort = new FakeExcelApi.RangeSort(this.context, _createPropertyObjectPath(this.context, this, "Sort", false, false));
                }
                return this.m_sort;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "columnIndex", {
            get: function () {
                _throwIfNotLoaded("columnIndex", this.m_columnIndex);
                return this.m_columnIndex;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "logText", {
            get: function () {
                _throwIfNotLoaded("logText", this.m_logText);
                return this.m_logText;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "rowIndex", {
            get: function () {
                _throwIfNotLoaded("rowIndex", this.m_rowIndex);
                return this.m_rowIndex;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "text", {
            get: function () {
                _throwIfNotLoaded("text", this.m_text);
                return this.m_text;
            },
            set: function (value) {
                this.m_text = value;
                _createSetPropertyAction(this.context, this, "Text", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "textArray", {
            get: function () {
                _throwIfNotLoaded("textArray", this.m_textArray);
                return this.m_textArray;
            },
            set: function (value) {
                this.m_textArray = value;
                _createSetPropertyAction(this.context, this, "TextArray", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "value", {
            get: function () {
                _throwIfNotLoaded("value", this.m_value);
                return this.m_value;
            },
            set: function (value) {
                this.m_value = value;
                _createSetPropertyAction(this.context, this, "Value", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "valueArray", {
            get: function () {
                _throwIfNotLoaded("valueArray", this.m_valueArray);
                return this.m_valueArray;
            },
            set: function (value) {
                this.m_valueArray = value;
                _createSetPropertyAction(this.context, this, "ValueArray", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "valueArray2", {
            get: function () {
                _throwIfNotLoaded("valueArray2", this.m_valueArray2);
                return this.m_valueArray2;
            },
            set: function (value) {
                this.m_valueArray2 = value;
                _createSetPropertyAction(this.context, this, "ValueArray2", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "valueTypes", {
            get: function () {
                _throwIfNotLoaded("valueTypes", this.m_valueTypes);
                return this.m_valueTypes;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(Range.prototype, "_ReferenceId", {
            get: function () {
                _throwIfNotLoaded("_ReferenceId", this.m__ReferenceId);
                return this.m__ReferenceId;
            },
            enumerable: true,
            configurable: true
        });
        Range.prototype.activate = function () {
            /* Begin_PlaceHolder_Range_Activate */
            /* End_PlaceHolder_Range_Activate */
            _createMethodAction(this.context, this, "Activate", 0 /* Default */, []);
        };
        Range.prototype.getValueArray2 = function () {
            /* Begin_PlaceHolder_Range_GetValueArray2 */
            /* End_PlaceHolder_Range_GetValueArray2 */
            var action = _createMethodAction(this.context, this, "GetValueArray2", 0 /* Default */, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        Range.prototype.notRestMethod = function () {
            /* Begin_PlaceHolder_Range_NotRestMethod */
            /* End_PlaceHolder_Range_NotRestMethod */
            var action = _createMethodAction(this.context, this, "NotRestMethod", 1 /* Read */, []);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        Range.prototype.replaceValue = function (newValue) {
            /* Begin_PlaceHolder_Range_ReplaceValue */
            if (console.log) {
                console.log("newValue" + newValue);
            }
            /* End_PlaceHolder_Range_ReplaceValue */
            var action = _createMethodAction(this.context, this, "ReplaceValue", 0 /* Default */, [newValue]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        Range.prototype.setValueArray2 = function (valueArray, text) {
            /* Begin_PlaceHolder_Range_SetValueArray2 */
            /* End_PlaceHolder_Range_SetValueArray2 */
            _createMethodAction(this.context, this, "SetValueArray2", 0 /* Default */, [valueArray, text]);
        };
        Range.prototype._KeepReference = function () {
            /* Begin_PlaceHolder_Range__KeepReference */
            /* End_PlaceHolder_Range__KeepReference */
            _createMethodAction(this.context, this, "_KeepReference", 1 /* Read */, []);
        };
        /** Handle results returned from the document
         * @private
         */
        Range.prototype._handleResult = function (value) {
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["ColumnIndex"])) {
                this.m_columnIndex = obj["ColumnIndex"];
            }
            if (!_isUndefined(obj["LogText"])) {
                this.m_logText = obj["LogText"];
            }
            if (!_isUndefined(obj["RowIndex"])) {
                this.m_rowIndex = obj["RowIndex"];
            }
            if (!_isUndefined(obj["Text"])) {
                this.m_text = obj["Text"];
            }
            if (!_isUndefined(obj["TextArray"])) {
                this.m_textArray = obj["TextArray"];
            }
            if (!_isUndefined(obj["Value"])) {
                this.m_value = obj["Value"];
            }
            if (!_isUndefined(obj["ValueArray"])) {
                this.m_valueArray = obj["ValueArray"];
            }
            if (!_isUndefined(obj["ValueArray2"])) {
                this.m_valueArray2 = obj["ValueArray2"];
            }
            if (!_isUndefined(obj["ValueTypes"])) {
                this.m_valueTypes = obj["ValueTypes"];
            }
            if (!_isUndefined(obj["_ReferenceId"])) {
                this.m__ReferenceId = obj["_ReferenceId"];
            }
            _handleNavigationPropertyResults(this, obj, ["sort", "Sort"]);
        };
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        Range.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        Range.prototype._initReferenceId = function (value) {
            this.m__ReferenceId = value;
        };
        return Range;
    })(OfficeExtension.ClientObject);
    FakeExcelApi.Range = Range;
    var TestCaseObject = (function (_super) {
        __extends(TestCaseObject, _super);
        function TestCaseObject() {
            _super.apply(this, arguments);
        }
        /* Begin_PlaceHolder_TestCaseObject_Custom_Members */
        /* End_PlaceHolder_TestCaseObject_Custom_Members */
        TestCaseObject.prototype.calculateAddressAndSaveToRange = function (street, city, range) {
            /* Begin_PlaceHolder_TestCaseObject_CalculateAddressAndSaveToRange */
            /* End_PlaceHolder_TestCaseObject_CalculateAddressAndSaveToRange */
            var action = _createMethodAction(this.context, this, "CalculateAddressAndSaveToRange", 0 /* Default */, [street, city, range]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        TestCaseObject.prototype.matrixSum = function (matrix) {
            /* Begin_PlaceHolder_TestCaseObject_MatrixSum */
            /* End_PlaceHolder_TestCaseObject_MatrixSum */
            var action = _createMethodAction(this.context, this, "MatrixSum", 0 /* Default */, [matrix]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        TestCaseObject.prototype.sum = function () {
            var values = [];
            for (var _i = 0; _i < arguments.length; _i++) {
                values[_i - 0] = arguments[_i];
            }
            /* Begin_PlaceHolder_TestCaseObject_Sum */
            /* End_PlaceHolder_TestCaseObject_Sum */
            var action = _createMethodAction(this.context, this, "Sum", 0 /* Default */, [values]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        TestCaseObject.prototype.testParamBool = function (value) {
            /* Begin_PlaceHolder_TestCaseObject_TestParamBool */
            /* End_PlaceHolder_TestCaseObject_TestParamBool */
            var action = _createMethodAction(this.context, this, "TestParamBool", 0 /* Default */, [value]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        TestCaseObject.prototype.testParamDouble = function (value) {
            /* Begin_PlaceHolder_TestCaseObject_TestParamDouble */
            /* End_PlaceHolder_TestCaseObject_TestParamDouble */
            var action = _createMethodAction(this.context, this, "TestParamDouble", 0 /* Default */, [value]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        TestCaseObject.prototype.testParamFloat = function (value) {
            /* Begin_PlaceHolder_TestCaseObject_TestParamFloat */
            /* End_PlaceHolder_TestCaseObject_TestParamFloat */
            var action = _createMethodAction(this.context, this, "TestParamFloat", 0 /* Default */, [value]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        TestCaseObject.prototype.testParamInt = function (value) {
            /* Begin_PlaceHolder_TestCaseObject_TestParamInt */
            /* End_PlaceHolder_TestCaseObject_TestParamInt */
            var action = _createMethodAction(this.context, this, "TestParamInt", 0 /* Default */, [value]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        TestCaseObject.prototype.testParamRange = function (value) {
            /* Begin_PlaceHolder_TestCaseObject_TestParamRange */
            /* End_PlaceHolder_TestCaseObject_TestParamRange */
            return new FakeExcelApi.Range(this.context, _createMethodObjectPath(this.context, this, "TestParamRange", 0 /* Default */, [value], false, false));
        };
        TestCaseObject.prototype.testParamString = function (value) {
            /* Begin_PlaceHolder_TestCaseObject_TestParamString */
            /* End_PlaceHolder_TestCaseObject_TestParamString */
            var action = _createMethodAction(this.context, this, "TestParamString", 0 /* Default */, [value]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        TestCaseObject.prototype.testUrlKeyValueDecode = function (value) {
            /* Begin_PlaceHolder_TestCaseObject_TestUrlKeyValueDecode */
            /* End_PlaceHolder_TestCaseObject_TestUrlKeyValueDecode */
            var action = _createMethodAction(this.context, this, "TestUrlKeyValueDecode", 0 /* Default */, [value]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        TestCaseObject.prototype.testUrlPathEncode = function (value) {
            /* Begin_PlaceHolder_TestCaseObject_TestUrlPathEncode */
            /* End_PlaceHolder_TestCaseObject_TestUrlPathEncode */
            var action = _createMethodAction(this.context, this, "TestUrlPathEncode", 0 /* Default */, [value]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        /** Handle results returned from the document
         * @private
         */
        TestCaseObject.prototype._handleResult = function (value) {
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
        };
        /**
         * Create a new instance of FakeExcelApi.TestCaseObject object
         */
        TestCaseObject.newObject = function (context) {
            var ret = new FakeExcelApi.TestCaseObject(context, _createNewObjectObjectPath(context, "ExcelApi.TestCaseObject", false));
            return ret;
        };
        return TestCaseObject;
    })(OfficeExtension.ClientObject);
    FakeExcelApi.TestCaseObject = TestCaseObject;
    var RangeSort = (function (_super) {
        __extends(RangeSort, _super);
        function RangeSort() {
            _super.apply(this, arguments);
        }
        Object.defineProperty(RangeSort.prototype, "fields", {
            /* Begin_PlaceHolder_RangeSort_Custom_Members */
            /* End_PlaceHolder_RangeSort_Custom_Members */
            get: function () {
                _throwIfNotLoaded("fields", this.m_fields);
                return this.m_fields;
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeSort.prototype, "fields2", {
            get: function () {
                _throwIfNotLoaded("fields2", this.m_fields2);
                return this.m_fields2;
            },
            set: function (value) {
                this.m_fields2 = value;
                _createSetPropertyAction(this.context, this, "Fields2", value);
            },
            enumerable: true,
            configurable: true
        });
        Object.defineProperty(RangeSort.prototype, "queryField", {
            get: function () {
                _throwIfNotLoaded("queryField", this.m_queryField);
                return this.m_queryField;
            },
            enumerable: true,
            configurable: true
        });
        RangeSort.prototype.apply = function (fields) {
            /* Begin_PlaceHolder_RangeSort_Apply */
            /* End_PlaceHolder_RangeSort_Apply */
            var action = _createMethodAction(this.context, this, "Apply", 0 /* Default */, [fields]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        RangeSort.prototype.applyAndReturnFirstField = function (fields) {
            /* Begin_PlaceHolder_RangeSort_ApplyAndReturnFirstField */
            /* End_PlaceHolder_RangeSort_ApplyAndReturnFirstField */
            var action = _createMethodAction(this.context, this, "ApplyAndReturnFirstField", 0 /* Default */, [fields]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        RangeSort.prototype.applyMixed = function (fields) {
            /* Begin_PlaceHolder_RangeSort_ApplyMixed */
            /* End_PlaceHolder_RangeSort_ApplyMixed */
            var action = _createMethodAction(this.context, this, "ApplyMixed", 0 /* Default */, [fields]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        RangeSort.prototype.applyQueryWithSortFieldAndReturnLast = function (fields) {
            /* Begin_PlaceHolder_RangeSort_ApplyQueryWithSortFieldAndReturnLast */
            /* End_PlaceHolder_RangeSort_ApplyQueryWithSortFieldAndReturnLast */
            var action = _createMethodAction(this.context, this, "ApplyQueryWithSortFieldAndReturnLast", 0 /* Default */, [fields]);
            var ret = new OfficeExtension.ClientResult();
            _addActionResultHandler(this, action, ret);
            return ret;
        };
        /** Handle results returned from the document
         * @private
         */
        RangeSort.prototype._handleResult = function (value) {
            if (_isNullOrUndefined(value))
                return;
            var obj = value;
            _fixObjectPathIfNecessary(this, obj);
            if (!_isUndefined(obj["Fields"])) {
                this.m_fields = obj["Fields"];
            }
            if (!_isUndefined(obj["Fields2"])) {
                this.m_fields2 = obj["Fields2"];
            }
            if (!_isUndefined(obj["QueryField"])) {
                this.m_queryField = obj["QueryField"];
            }
        };
        /**
         * Queues up a command to load the specified properties of the object. You must call "context.sync()" before reading the properties.
         */
        RangeSort.prototype.load = function (option) {
            _load(this, option);
            return this;
        };
        return RangeSort;
    })(OfficeExtension.ClientObject);
    FakeExcelApi.RangeSort = RangeSort;
    var ErrorCodes;
    (function (ErrorCodes) {
        ErrorCodes.aborted2 = "Aborted2";
        ErrorCodes.accessDenied = "AccessDenied";
        ErrorCodes.accessDenied2 = "AccessDenied2";
        ErrorCodes.conflict = "Conflict";
        ErrorCodes.conflict2 = "Conflict2";
        ErrorCodes.generalException = "GeneralException";
        ErrorCodes.invalidArgument = "InvalidArgument";
        ErrorCodes.outOfRange = "OutOfRange";
    })(ErrorCodes = FakeExcelApi.ErrorCodes || (FakeExcelApi.ErrorCodes = {}));
})(FakeExcelApi || (FakeExcelApi = {})); // FakeExcelApi
/* Begin_PlaceHolder_GlobalFooter */
/* End_PlaceHolder_GlobalFooter */
var FakeExcelApi;
(function (FakeExcelApi) {
    var ExcelClientRequestContext = (function (_super) {
        __extends(ExcelClientRequestContext, _super);
        function ExcelClientRequestContext(url) {
            _super.call(this, url);
            if (window.document && window.document.getElementById("RichApiFakeXlapiUseWac") && window.document.getElementById("RichApiFakeXlapiUseWac").checked) {
                this._requestExecutor = new OfficeExtensionTest.PostMessageRequestExecutor();
            }
            else {
                this._requestExecutor = new OfficeExtensionTest.InProcRequestExecutor();
            }
            this.m_application = new FakeExcelApi.Application(this, OfficeExtension.ObjectPathFactory.createGlobalObjectObjectPath(this));
            this._rootObject = this.m_application;
        }
        Object.defineProperty(ExcelClientRequestContext.prototype, "application", {
            get: function () {
                if (this.m_application == null) {
                    this.m_application = new FakeExcelApi.Application(this, OfficeExtension.ObjectPathFactory.createGlobalObjectObjectPath(this));
                }
                return this.m_application;
            },
            enumerable: true,
            configurable: true
        });
        return ExcelClientRequestContext;
    })(OfficeExtension.ClientRequestContext);
    FakeExcelApi.ExcelClientRequestContext = ExcelClientRequestContext;
    function run(batch) {
        return OfficeExtension.ClientRequestContext._run(function () { return new FakeExcelApi.ExcelClientRequestContext(); }, batch);
    }
    FakeExcelApi.run = run;
})(FakeExcelApi || (FakeExcelApi = {}));
var OfficeExtensionTest;
(function (OfficeExtensionTest) {
    function fixupResponse(result) {
        for (var i = 0; i < result.length; i++) {
            if (result[i] != null && result[i].toArray) {
                result[i] = result[i].toArray();
            }
        }
    }
    var InProcRequestExecutor = (function () {
        function InProcRequestExecutor() {
        }
        InProcRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage, callback) {
            var requestMessageText = JSON.stringify(requestMessage.Body);
            OfficeExtension.Utility.log("Request:");
            OfficeExtension.Utility.log(requestMessageText);
            var messageSafearray = OfficeExtension.RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, "POST", "ProcessQuery", null, requestMessageText);
            OfficeExtension.Utility.log("RequestSafeArray:");
            OfficeExtension.Utility.log(JSON.stringify(messageSafearray));
            window.setTimeout(function () {
                var ext = window["RichApiExecutorTestControl"];
                var response = { ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
                var responseVBArray = ext.Execute(messageSafearray);
                var result = new VBArray(responseVBArray).toArray();
                fixupResponse(result);
                OfficeExtension.Utility.log("Response:");
                OfficeExtension.Utility.log(JSON.stringify(result));
                var bodyText = OfficeExtension.RichApiMessageUtility.getResponseBodyFromSafeArray(result);
                response.Body = JSON.parse(bodyText);
                response.Headers = OfficeExtension.RichApiMessageUtility.getResponseHeadersFromSafeArray(result);
                callback(response);
            }, 100);
        };
        return InProcRequestExecutor;
    })();
    OfficeExtensionTest.InProcRequestExecutor = InProcRequestExecutor;
    function inProcExecuteRichApiRequestAsync(messageSafearray, callback) {
        window.setTimeout(function () {
            var ext = window["RichApiExecutorTestControl"];
            var response = { status: "succeeded", value: { data: [] }, error: null };
            var responseVBArray = ext.Execute(messageSafearray);
            var result = new VBArray(responseVBArray).toArray();
            fixupResponse(result);
            OfficeExtension.Utility.log("Response:");
            OfficeExtension.Utility.log(JSON.stringify(result));
            response.value.data = result;
            callback(response);
        }, 100);
    }
    OfficeExtensionTest.inProcExecuteRichApiRequestAsync = inProcExecuteRichApiRequestAsync;
})(OfficeExtensionTest || (OfficeExtensionTest = {}));
var OfficeExtensionTest;
(function (OfficeExtensionTest) {
    var PostMessageRequestExecutor = (function () {
        function PostMessageRequestExecutor() {
        }
        PostMessageRequestExecutor.ensureInit = function () {
            if (!PostMessageRequestExecutor.s_callbackMap) {
                PostMessageRequestExecutor.s_callbackMap = {};
                window.addEventListener("message", function (ev) {
                    var msg = JSON.parse(ev.data);
                    if (msg && msg.messageId && msg.messageType == 'RichApiResponse' && msg.message) {
                        var callback = PostMessageRequestExecutor.s_callbackMap[msg.messageId];
                        if (callback) {
                            delete PostMessageRequestExecutor.s_callbackMap[msg.messageId];
                            var response = { ErrorCode: '', ErrorMessage: '', Headers: null, Body: null };
                            OfficeExtension.Utility.log("Response:");
                            OfficeExtension.Utility.log(JSON.stringify(msg.message));
                            var bodyText = OfficeExtension.RichApiMessageUtility.getResponseBodyFromSafeArray(msg.message);
                            response.Body = JSON.parse(bodyText);
                            response.Headers = OfficeExtension.RichApiMessageUtility.getResponseHeadersFromSafeArray(msg.message);
                            callback(response);
                        }
                    }
                });
            }
        };
        PostMessageRequestExecutor.prototype.executeAsync = function (customData, requestFlags, requestMessage, callback) {
            PostMessageRequestExecutor.ensureInit();
            var requestMessageText = JSON.stringify(requestMessage.Body);
            OfficeExtension.Utility.log("Request:");
            OfficeExtension.Utility.log(requestMessageText);
            var messageSafearray = OfficeExtension.RichApiMessageUtility.buildRequestMessageSafeArray(customData, requestFlags, "POST", "ProcessQuery", null, requestMessageText);
            OfficeExtension.Utility.log("RequestSafeArray:");
            OfficeExtension.Utility.log(JSON.stringify(messageSafearray));
            var msgId = PostMessageRequestExecutor.s_messageId;
            PostMessageRequestExecutor.s_messageId = PostMessageRequestExecutor.s_messageId + 1;
            var msg = { messageId: msgId, messageType: 'RichApiRequest', message: messageSafearray };
            var frame = document.getElementById(PostMessageRequestExecutor.s_frameId);
            PostMessageRequestExecutor.s_callbackMap[msgId] = callback;
            frame.contentWindow.postMessage(JSON.stringify(msg), "*");
        };
        PostMessageRequestExecutor.s_messageId = 1;
        PostMessageRequestExecutor.s_frameId = "RichApiWacFrame";
        return PostMessageRequestExecutor;
    })();
    OfficeExtensionTest.PostMessageRequestExecutor = PostMessageRequestExecutor;
})(OfficeExtensionTest || (OfficeExtensionTest = {}));

var ExcelTest;
(function (ExcelTest) {
    function reportError(errorInfo) {
        if (errorInfo instanceof OfficeExtension.Error) {
            RichApiTest.log.comment("ErrorCode=" + errorInfo.code);
            RichApiTest.log.comment("ErrorLocation=" + errorInfo.debugInfo.errorLocation);
            RichApiTest.log.comment("ErrorMessage=" + errorInfo.message);
        }
        else if (errorInfo.code && errorInfo.message) {
            RichApiTest.log.comment("ErrorCode=" + errorInfo.code);
            RichApiTest.log.comment("ErrorMessage=" + errorInfo.message);
        }
        else {
            RichApiTest.log.comment("Error=" + JSON.stringify(errorInfo));
        }
        RichApiTest.log.done(false);
    }
    ExcelTest.reportError = reportError;
    /**
        For use as the final .then clause of a promise, before the .catch.  E.g.,: ".then(pass).catch(ExcelTest.reportError)"
    */
    function pass() {
        RichApiTest.log.done(true);
    }
    ExcelTest.pass = pass;
    function reportJQueryError(xhr) {
        RichApiTest.log.comment("StatusCode=" + xhr.status);
        RichApiTest.log.comment("ResponseText=" + xhr.responseText);
        RichApiTest.log.done(false);
    }
    ExcelTest.reportJQueryError = reportJQueryError;
    var Settings = (function () {
        function Settings() {
            this.baseUri = OfficeExtension.Constants.localDocumentApiPrefix;
        }
        return Settings;
    })();
    ExcelTest.Settings = Settings;
    ExcelTest.settings = new Settings();
    function getNewScriptText() {
        return [
            'var ctx = new Excel.RequestContext();',
            '',
            'var sheetName = "Sheet1";',
            'var worksheet = ctx.workbook.worksheets.getItem(sheetName);',
            'ctx.load(worksheet);',
            '',
            'ctx.sync()',
            '    .then(function () {',
            '        RichApiTest.log.comment("");',
            '    })',
            '    .then(ExcelTest.pass)',
            '    .catch(ExcelTest.reportError);'
        ].join('\n');
    }
    ExcelTest.getNewScriptText = getNewScriptText;
})(ExcelTest || (ExcelTest = {}));
/**
    This RestUtility wuill eventually be merged into RichApiTest.Core in OSFClient, once we're done writing tests.
    But for now, due to branching complexity, it's easier to accumulate everything here, then copy over to OSFClient in one go,
    and then remove the duplication.
*/
var ExcelTest;
(function (ExcelTest) {
    var RestUtility;
    (function (RestUtility) {
        (function (Status) {
            Status[Status["OK"] = 200] = "OK";
            Status[Status["Created"] = 201] = "Created";
            Status[Status["NoContent"] = 204] = "NoContent";
            Status[Status["BadRequest"] = 400] = "BadRequest";
            Status[Status["NotFound"] = 404] = "NotFound";
            Status[Status["MethodNotAllowed"] = 405] = "MethodNotAllowed";
        })(RestUtility.Status || (RestUtility.Status = {}));
        var Status = RestUtility.Status;
        var Thenable;
        (function (Thenable) {
            /**
             * Posts a request with a message body (e.g., posting to collection)
             * @param relativePath - path of the action relative to the root, like "worksheets('charts')/charts"
             * @param body - an object that will get JSON-stringified into the post message body.
            */
            function post(relativePath, body) {
                return function () { return RestUtility.post(relativePath, body); };
            }
            Thenable.post = post;
            /**
             * Posts a method action to a relative URL path.
             * @param relativePath - path of the action relative to the root, like "worksheets('charts')/charts/$/add"
             * @param mathodParameters - an object literal (not string) with parameter key-value pairs (e.g., "{ 'type': 'pie', 'sourceData': 'A1:B4' }")
             */
            function postAsUrlAction(relativePath, methodParameters) {
                return function () { return RestUtility.postAsUrlAction(relativePath, methodParameters); };
            }
            Thenable.postAsUrlAction = postAsUrlAction;
            /** Performs a REST get based on a path */
            function get(relativePath) {
                return function () { return RestUtility.get(relativePath); };
            }
            Thenable.get = get;
            /** performs a PATCH operation on an item, passing in the specified properties */
            function patch(id, properties) {
                return function () { return RestUtility.patch(id, properties); };
            }
            Thenable.patch = patch;
            /** Gets an object of type T out of the response. Optionally applies a transformation function (e.g., to get a particular property out of the body). */
            function getBodyAsObject(transform) {
                return function (resp) { return RestUtility.getBodyAsObject(resp, transform); };
            }
            Thenable.getBodyAsObject = getBodyAsObject;
            /** Verifies that the object type matches the expected type, then pipe the object through */
            function verifyObjectType(expectedObjectType) {
                return function (obj) {
                    var odataType = getODataType(obj);
                    var same = odataType === expectedObjectType;
                    if (!same && odataType && expectedObjectType) {
                        var index;
                        index = odataType.lastIndexOf(".");
                        if (index > 0) {
                            odataType = odataType.substr(index + 1);
                        }
                        index = expectedObjectType.lastIndexOf(".");
                        if (index > 0) {
                            expectedObjectType = expectedObjectType.substr(index + 1);
                        }
                        same = odataType === expectedObjectType;
                    }
                    Util.assert(same);
                    return obj;
                };
            }
            Thenable.verifyObjectType = verifyObjectType;
            /** Validates the status code of a response (throwing on invalid), and then pipes the RestResponseInfo to the next .then */
            function validateStatus(expected) {
                return function (resp) {
                    RichApiTest.RestUtility.verifyStatusCodeThrow(resp, expected);
                    return resp;
                };
            }
            Thenable.validateStatus = validateStatus;
            /** Validates the status code of an array of responses (throwing on invalid), and then pipes the RestResponseInfo's to the next .then */
            function validateStatuses(expected) {
                return function (responses) { return responses.map(validateStatus(expected)); };
            }
            Thenable.validateStatuses = validateStatuses;
            function validateErrorCode(expectedCode) {
                return function (resp) {
                    if (OfficeExtension.Utility.isNullOrEmptyString(resp.body)) {
                        throw new Error("Empty body");
                    }
                    var bodyTrimmed = OfficeExtension.Utility.trim(resp.body);
                    if (bodyTrimmed.length == 0) {
                        throw new Error("Empty body");
                    }
                    if (bodyTrimmed.charAt(0) != '{') {
                        throw new Error("Not JSON body:" + resp.body);
                    }
                    var errorBody = JSON.parse(resp.body);
                    if (!errorBody) {
                        throw new Error("Cannot parse response body: " + resp.body);
                    }
                    if (!errorBody.error) {
                        throw new Error("Cannot get error from response body: " + resp.body);
                    }
                    Util.assertCompareValues(expectedCode, errorBody.error.code, "Error response code");
                    return resp;
                };
            }
            Thenable.validateErrorCode = validateErrorCode;
            /** invoke get method for given url and validate response using promise, and then pipes void to the next .then */
            function invokeGetAndValidateResponsePromise(relativeUrl, expectedValue, transform) {
                return function () {
                    return RestUtility.invokeGetAndValidateResponsePromise(relativeUrl, expectedValue, transform);
                };
            }
            Thenable.invokeGetAndValidateResponsePromise = invokeGetAndValidateResponsePromise;
        })(Thenable = RestUtility.Thenable || (RestUtility.Thenable = {}));
        /**
         * invoke get method for given url and validate response using promise
         * @param relativePath - path of the action relative to the root, like "worksheets('sheet1')".
         * @param expectedValue - expected value in response.
         * @param transform - transform function.
         */
        function invokeGetAndValidateResponsePromise(relativeUrl, expectedValue, transform) {
            return RestUtility.get(relativeUrl).then(RestUtility.Thenable.validateStatus(200 /* OK */)).then(RestUtility.Thenable.getBodyAsObject(transform)).then(function (value) {
                RestUtility.validateValue(value, expectedValue);
            });
        }
        RestUtility.invokeGetAndValidateResponsePromise = invokeGetAndValidateResponsePromise;
        /**
         * invoke get method for given url and validate response
         * @param relativePath - path of the action relative to the root, like "worksheets('sheet1')".
         * @param expectedValue - expected value in response.
         * @param transform - transform function.
         */
        function invokeGetAndValidateResponse(relativeUrl, expectedValue, transform) {
            RestUtility.invokeGetAndValidateResponsePromise(relativeUrl, expectedValue, transform).then(ExcelTest.pass).catch(ExcelTest.reportError);
        }
        RestUtility.invokeGetAndValidateResponse = invokeGetAndValidateResponse;
        /**
         * Validates that the value and expected value are same with each other
         */
        function validateValue(value, expectedValue, delta) {
            Util.assert(Util.compare(value, expectedValue, delta));
        }
        RestUtility.validateValue = validateValue;
        /**
         * Ensure the disable API fails with expected error
         * @param relativePath - path of the action relative to the root, like "worksheets('sheet1')".
         * @param httpMethod - REST request method.
         */
        function ensureDisabledApiFailed(relativePath, httpMethod) {
            var expectedErrorCode = Excel.ErrorCodes.invalidArgument;
            var expectedErrorCodeFromGraph = "BadRequest";
            RichApiTest.RestUtility.invoke({
                method: httpMethod,
                url: RichApiTest.RestUtility.getBaseUrlUsingOverride(ExcelTest.settings.baseUri) + relativePath,
                body: ""
            }).then(function (resp) {
                RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusBadRequest);
                var responseJson = JSON.parse(resp.body);
                if (responseJson) {
                    if (responseJson.error.code != expectedErrorCode && responseJson.error.code != expectedErrorCodeFromGraph) {
                        throw new Error("Expected response code is " + expectedErrorCode + "; actual response code is " + responseJson.error.code);
                    }
                }
                else {
                    throw new Error("Cannot parse response");
                }
            }).then(ExcelTest.pass).catch(ExcelTest.reportError);
        }
        RestUtility.ensureDisabledApiFailed = ensureDisabledApiFailed;
        /**
         * Posts a request with a message body (e.g., posting to collection)
         * @param relativePath - path of the action relative to the root, like "worksheets('charts')/charts"
         * @param body - an object that will get JSON-stringified into the post message body.
         */
        function post(relativePath, body) {
            var request = {
                url: RichApiTest.RestUtility.getBaseUrlUsingOverride(ExcelTest.settings.baseUri) + relativePath,
                method: RichApiTest.RestUtility.httpMethodPost,
                body: JSON.stringify(body)
            };
            RichApiTest.log.comment("Invoking POST with URL: " + request.url);
            return RichApiTest.RestUtility.invoke(request);
        }
        RestUtility.post = post;
        /**
         * Posts a method action to a relative URL path.
         * @param relativePath - path of the action relative to the root, like "worksheets('charts')/charts/$/add"
         * @param mathodParameters (optional) - an object literal (not string) with parameter key-value pairs (e.g., "{ 'type': 'pie', 'sourceData': 'A1:B4' }")
         */
        function postAsUrlAction(relativePath, methodParameters) {
            var body = "";
            if (methodParameters) {
                body = JSON.stringify(methodParameters);
            }
            var request = {
                url: RichApiTest.RestUtility.getBaseUrlUsingOverride(ExcelTest.settings.baseUri) + relativePath,
                body: body,
                method: RichApiTest.RestUtility.httpMethodPost,
            };
            RichApiTest.log.comment("Invoking POST method action with URL: " + request.url);
            return RichApiTest.RestUtility.invoke(request);
        }
        RestUtility.postAsUrlAction = postAsUrlAction;
        /** Performs a REST get based on a path */
        function get(relativePath) {
            var request = {
                url: RichApiTest.RestUtility.getBaseUrlUsingOverride(ExcelTest.settings.baseUri) + relativePath,
                method: RichApiTest.RestUtility.httpMethodGet
            };
            RichApiTest.log.comment("Invoking GET with URL: " + request.url);
            return RichApiTest.RestUtility.invoke(request);
        }
        RestUtility.get = get;
        /** performs a PATCH operation on an item, passing in the specified properties */
        function patch(id, properties) {
            var parsedBody = JSON.stringify(properties);
            var request = {
                url: RichApiTest.RestUtility.getBaseUrlUsingOverride(ExcelTest.settings.baseUri) + id,
                method: RichApiTest.RestUtility.httpMethodPatch,
                body: parsedBody
            };
            RichApiTest.log.comment("Invoking PATCH with URL: " + request.url);
            RichApiTest.log.comment("Invoking PATCH with Body: " + parsedBody);
            return RichApiTest.RestUtility.invoke(request);
        }
        RestUtility.patch = patch;
        /** Deletes an item either based on its "@odata.id" field. Can pass in either just the ID string, or a full object that contains this field */
        function deleteItem(obj) {
            var relativePath = extractId(obj);
            var request = {
                url: RichApiTest.RestUtility.getBaseUrlUsingOverride(ExcelTest.settings.baseUri) + relativePath,
                method: RichApiTest.RestUtility.httpMethodDelete
            };
            RichApiTest.log.comment("Invoking DELETE with URL: " + request.url);
            return RichApiTest.RestUtility.invoke(request);
        }
        RestUtility.deleteItem = deleteItem;
        /** Gets an object of type T out of the response. Optionally applies a transformation function (e.g., to get a particular property out of the body). */
        function getBodyAsObject(resp, transform) {
            var obj = JSON.parse(resp.body);
            if (transform) {
                obj = transform(obj);
            }
            return obj;
        }
        RestUtility.getBodyAsObject = getBodyAsObject;
        /** Gets the "@odata.id" property from the object */
        function getODataId(obj) {
            var ret = decodeURI(obj["@odata.id"]);
            if (ret) {
                var index = ret.toLowerCase().indexOf('workbook/');
                if (index >= 0) {
                    ret = ret.substr(index + 'workbook/'.length);
                }
            }
            return ret;
        }
        RestUtility.getODataId = getODataId;
        /** Gets the "@odata.id" property from the object */
        function getODataType(obj) {
            return obj["@odata.type"];
        }
        RestUtility.getODataType = getODataType;
        /** Extracts the "@odata.id" property from the object, unless the passed-in parameter is already a string -- in which case just use it as is */
        function extractId(obj) {
            if (typeof obj === 'string') {
                return obj;
            }
            var extracted = getODataId(obj);
            if (extracted != undefined) {
                return extracted;
            }
            throw new Error("Cannot get the OData id for the object. The parameter is neither a string, nor an object that has a '@odata.id' field.");
        }
        /**
            Encodes an odata literal (e.g., when a parameter is passed in to a method, as in:
            worksheets('charts')/charts/$/add(type='pie',sourceData='a1:b4')
    
            While the encoding is not necessary for the simple case above, if instead of 'a1:b4' the text contained
            an apostrope or invalid URL characters, it would need to be appropriately encoded.
        */
        function encodeODataLiteral(value) {
            if (typeof value === 'string') {
                // Replace single apostrophe with double-apostrophe
                return "'" + replaceAll(value, "'", "''") + "'";
            }
            else if (typeof value == 'boolean' || typeof value == 'number') {
                return value;
            }
            else {
                throw new Error("Unsupported type " + typeof value);
            }
        }
        RestUtility.encodeODataLiteral = encodeODataLiteral;
        /** Search-and-replace utility (JavaScript's string.replace only replaces the first occurence) */
        function replaceAll(text, search, replace) {
            return text.split(search).join(replace);
        }
    })(RestUtility = ExcelTest.RestUtility || (ExcelTest.RestUtility = {}));
})(ExcelTest || (ExcelTest = {}));
var Util;
(function (Util) {
    /**
     * Calls sync and ensures that the execution failed as expected
     * @param context - Request context
     * @param {string} expectedErrorCode - Expected error code
     * @param additionalValidation - (Optional) Additional error validation function, if any (e.g., to validate trace messages)
     * @param onCompletion - (Optional) Function callback for what to do when the expected failure was received.
           If left empty, the code will simply call RichApiTest.log.done(true) to signal the end-of-test.
     */
    function ensureSyncFailed(ctx, expectedErrorCode, additionalValidation, onCompletion) {
        ensureSyncFailedPromise(ctx, expectedErrorCode).then(function (error) {
            // Is there somethign else we need to validate?
            if (additionalValidation != null) {
                if (!additionalValidation(error)) {
                    RichApiTest.log.fail("... However, the additional validation failed.");
                    return;
                }
            }
            if (onCompletion != null) {
                onCompletion();
            }
            else {
                // Otherwise, just assume that wanted to exit the test as "succeeded"
                RichApiTest.log.done(true);
            }
        }).catch(function (error) {
            RichApiTest.log.fail(error);
        });
    }
    Util.ensureSyncFailed = ensureSyncFailed;
    function ensureSyncFailedPromise(ctx, expectedErrorCode) {
        return ctx.sync().then(function () {
            throw new Error("Sync should have throw, so this shouldn't have been reached.");
            return null; // Need this so that TypeScript compiler sees that the function returns a value (throwing is apparently not enough)
        }, function (error) {
            if (error.code == expectedErrorCode) {
                RichApiTest.log.comment("Got expected failure: " + error.code + ": " + error.message);
                RichApiTest.log.comment("ErrorLocation: " + error.debugInfo.errorLocation);
            }
            else {
                RichApiTest.log.comment("Received failure, but not the expected one!");
                RichApiTest.log.comment("Expected error code was " + expectedErrorCode);
                throw new Error("Received error was " + error.code + ": " + error.message);
            }
            return error;
        });
    }
    Util.ensureSyncFailedPromise = ensureSyncFailedPromise;
    function promisify(action) {
        return new OfficeExtension.Promise(function (resolve, reject) {
            var callback = function (result) {
                if (result.status == Office.AsyncResultStatus.Failed) {
                    reject(result.error);
                }
                else {
                    resolve(result.value);
                }
            };
            action(callback);
        });
    }
    Util.promisify = promisify;
    function throwOfficeError(error) {
        throw new Error(error.code + ":" + error.message);
    }
    Util.throwOfficeError = throwOfficeError;
    function ensureExpectedFailure(e, expectedErrorCode) {
        if ((e instanceof OfficeExtension.Error) && (e.code = expectedErrorCode)) {
            RichApiTest.log.pass("Caught expected exception");
        }
        else {
            RichApiTest.log.fail("Caught exception, but not of the expected type. Expecting " + expectedErrorCode + ", received " + JSON.stringify(e));
        }
    }
    Util.ensureExpectedFailure = ensureExpectedFailure;
    function wait(milliseconds, action) {
        var promise = new OfficeExtension.Promise(function (resolve, reject) {
            setTimeout(function () {
                resolve();
            }, milliseconds);
        });
        if (action) {
            return promise.then(action);
        }
        else {
            return promise;
        }
    }
    Util.wait = wait;
    function moveSheet(sheetName, target) {
        var ctx = new Excel.RequestContext();
        var sheets = ctx.workbook.worksheets;
        var sheet = ctx.workbook.worksheets.getItem(sheetName);
        var expected = target;
        ctx.load(sheets);
        ctx.load(sheet);
        ctx.sync().then(function () {
            var position = sheet.position;
            if (expected == -1) {
                expected = sheets.items.length;
            }
            sheet.position = target;
            ctx.load(sheet);
            ctx.sync().then(function () {
                var success = true;
                if (sheet.position != expected) {
                    RichApiTest.log.comment("Sheet: " + sheet.name);
                    RichApiTest.log.comment("position: " + position);
                    RichApiTest.log.comment("Expect sheet at: " + expected);
                    RichApiTest.log.comment("Actual sheet at: " + sheet.position);
                    success = false;
                }
                else {
                    RichApiTest.log.comment("Sheet: " + sheet.name);
                    RichApiTest.log.comment("Moved to: " + sheet.position);
                }
                sheet.position = position; // restore it.
                ctx.load(sheet);
                ctx.sync().then(function () {
                    if (sheet.position != position) {
                        RichApiTest.log.comment("moved to: " + sheet.position);
                        RichApiTest.log.comment("should be: " + position);
                        RichApiTest.log.comment("Restore failed");
                        success = false;
                    }
                    RichApiTest.log.done(success);
                }, ExcelTest.reportError);
            }, ExcelTest.reportError);
        }, ExcelTest.reportError);
    }
    Util.moveSheet = moveSheet;
    /**
        Removes all charts from the sheet, and then either calls the passed-in test function, or just returns a promise that can be chained on to another action.
    */
    function removeAllChartsBefore(sheetName, test) {
        RichApiTest.log.comment('Cleaning all charts from the sheet "' + sheetName + '"');
        var ctx = new Excel.RequestContext();
        var charts = ctx.workbook.worksheets.getItem(sheetName).charts;
        ctx.load(charts, "id");
        var result = ctx.sync().then(function () {
            RichApiTest.log.comment(charts.count + " charts were found");
            for (var i = 0; i < charts.count; i++) {
                charts.getItemAt(0).delete();
            }
        }).then(ctx.sync).then(function () { return RichApiTest.log.comment("Charts cleaned, running the requested action..."); });
        if (test) {
            result = result.then(function () {
                // This is for old-style tets that didn't use promises. Just invoke the test as a void function
                //    (and have the catch at the bottom only to catch the errors from upstream, during chart deletion)
                test();
            }).catch(ExcelTest.reportError);
            // If ran callback test, then value of promise isn't meant to be used -- wouldn't chain correctly anyway
            return null;
        }
        else {
            // Return the promise having not handled any errors -- those will be handled downstream by the function consumer anyway
            return result;
        }
    }
    Util.removeAllChartsBefore = removeAllChartsBefore;
    function clearSheetBefore(sheetName, test) {
        RichApiTest.log.comment('Clearing Sheet "' + sheetName + '"');
        var ctx = new Excel.RequestContext();
        var rangeClear = ctx.workbook.worksheets.getItem(sheetName).getRange(null);
        rangeClear.clear(null);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Sheet cleared, running the requested action...");
            test();
        });
    }
    Util.clearSheetBefore = clearSheetBefore;
    function clearSheetRest(sheetName) {
        RichApiTest.log.comment('Clearing Sheet "' + sheetName + '"');
        // OfficeMain: 2899699: When there is a table in the sheet, call sheet.range(null).clear(null) will get access violation
        // We need to delete the table first before call sheet.range(null).clear()
        return ExcelTest.RestUtility.get("worksheets('" + sheetName + "')/tables").then(ExcelTest.RestUtility.Thenable.validateStatus(200 /* OK */)).then(ExcelTest.RestUtility.Thenable.getBodyAsObject(function (obj) { return obj.value; })).then(function (result) { return OfficeExtension.Promise.all(result.map(function (item) { return ExcelTest.RestUtility.deleteItem(item); })); }).then(ExcelTest.RestUtility.Thenable.validateStatuses(204 /* NoContent */)).then(ExcelTest.RestUtility.Thenable.post("worksheets('" + sheetName + "')/range(address=null)/clear", { applyTo: Excel.ClearApplyTo.all })).then(ExcelTest.RestUtility.Thenable.validateStatus(204 /* NoContent */)).then(function () {
            RichApiTest.log.comment("Sheet cleared, running the requested action...");
        });
    }
    Util.clearSheetRest = clearSheetRest;
    function checkSheetExistsRunTest(ctx, sheetName, onCompletion) {
        var sheet = ctx.workbook.worksheets.getItem(sheetName);
        ctx.load(sheet);
        ctx.sync().then(function () {
            RichApiTest.log.comment("Sheet '" + sheetName + "' was found, proceeding with test execution.");
            if (onCompletion != null) {
                onCompletion();
            }
            else {
                // Otherwise, just assume that wanted to exit the test as "succeeded"
                RichApiTest.log.done(true);
            }
        }, function (errorInfo) {
            if (errorInfo.code == Excel.ErrorCodes.itemNotFound) {
                RichApiTest.log.comment("Sheet '" + sheetName + "' was not found, skipping test.");
                RichApiTest.log.done(true);
            }
            else {
                RichApiTest.log.comment("Received failure, but not the expected one!");
                RichApiTest.log.fail("Received error was " + errorInfo.code + ": " + errorInfo.message);
            }
        });
    }
    Util.checkSheetExistsRunTest = checkSheetExistsRunTest;
    function getTableAtPosition(address, callback) {
        var ctx = new Excel.RequestContext();
        var tables = ctx.workbook.tables;
        ctx.load(tables, "id");
        ctx.sync().then(function () {
            var tableVisited = 0;
            for (var i = 0; i < tables.count; i++) {
                var table = tables.items[i];
                var range = table.getRange();
                ctx.load(range, "address");
                (function (boundTable, boundRange) {
                    ctx.sync().then(function () {
                        if (boundRange.address == address) {
                            callback(boundTable);
                        }
                        else {
                            tableVisited++;
                            if (tableVisited == tables.count) {
                                callback(null);
                            }
                        }
                    }, ExcelTest.reportError);
                })(table, range);
            }
        }, ExcelTest.reportError);
    }
    Util.getTableAtPosition = getTableAtPosition;
    function assert(statementOrLambda, explanation) {
        if (typeof (statementOrLambda) === "boolean") {
            if (!statementOrLambda) {
                throw new Error("Assert failed" + (explanation ? (" because " + explanation) : ""));
            }
            else {
                RichApiTest.log.comment("Assert passed" + (explanation ? (": " + explanation) : ""));
            }
        }
        else {
            assert(statementOrLambda(), statementOrLambda.toString());
        }
    }
    Util.assert = assert;
    function assertCompareArray(value, expected) {
        if (!compareArray(value, expected)) {
            throw new Error("Arrays did not match");
        }
        else {
            RichApiTest.log.comment("Arrays are matching.");
        }
    }
    Util.assertCompareArray = assertCompareArray;
    function assertCompareValues(expectedValue, actualValue, additionalComment) {
        var prefix = (additionalComment ? additionalComment + " - " : "");
        assert(actualValue === expectedValue, prefix + "Expected: '" + expectedValue + "' Actual: '" + actualValue + "'");
    }
    Util.assertCompareValues = assertCompareValues;
    function assertCompareNumeric(expectedValue, actualValue, allowableMarginOfError, additionalComment) {
        var prefix = (additionalComment ? additionalComment + " - " : "");
        assert(Math.abs(actualValue - expectedValue) <= allowableMarginOfError, prefix + "Expected: '" + expectedValue + "' Actual: '" + actualValue + "'");
    }
    Util.assertCompareNumeric = assertCompareNumeric;
    function compare(value, expected, delta) {
        var result = (value === expected);
        if (result) {
            RichApiTest.log.comment('Compare succeeded: expected "' + expected + '" and received it.');
        }
        else {
            if (typeof (delta) != "undefined" && typeof (value) == "number" && typeof (expected) == "number" && Math.abs(value - expected) < delta) {
                result = true;
                RichApiTest.log.comment('Compare succeeded: expected "' + expected + '" and received it.');
            }
            else {
                RichApiTest.log.comment('Compare FAILED: expected "' + expected + '", received "' + value + '"');
            }
        }
        return result;
    }
    Util.compare = compare;
    function parseStringAsNumberAndCompare(value, expected) {
        if (typeof (value == "string")) {
            value = parseFloat(value);
        }
        return compare(value, expected);
    }
    Util.parseStringAsNumberAndCompare = parseStringAsNumberAndCompare;
    function compareProperty(object, propertyName, expected) {
        var result = (object[propertyName] === expected);
        if (!result) {
            RichApiTest.log.fail(propertyName + " was set to: '" + object[propertyName] + "'. Expected :'" + expected + "'.");
        }
        return result;
    }
    Util.compareProperty = compareProperty;
    function compareArray(value, expected) {
        if (!expected || !value) {
            return false;
        }
        if (value.length != expected.length) {
            RichApiTest.log.comment("Expected array length: " + expected.length + " Actual: " + value.length);
            return false;
        }
        for (var i = 0, l = value.length; i < l; i++) {
            if (value[i] instanceof Array && expected[i] instanceof Array) {
                if (!compareArray(value[i], expected[i])) {
                    return false;
                }
            }
            else if (value[i] != expected[i]) {
                RichApiTest.log.comment("Expected value: " + expected[i] + " Actual: " + value[i]);
                return false;
            }
        }
        return true;
    }
    Util.compareArray = compareArray;
    function startsWith(thisString, searchString, position) {
        position = position || 0;
        return thisString.substr(position, searchString.length) === searchString;
    }
    Util.startsWith = startsWith;
    function isWAC() {
        return window.location.href.indexOf("|Web") >= 0;
    }
    Util.isWAC = isWAC;
    function isDesktop() {
        return window.location.href.indexOf("|Web") == -1;
    }
    Util.isDesktop = isDesktop;
    function isMacOS() {
        return window.navigator.platform.toUpperCase().indexOf("MAC") >= 0;
    }
    Util.isMacOS = isMacOS;
    function isiOS() {
        return window.navigator.userAgent.match(/(iPhone|iPad|iPod)/i);
    }
    Util.isiOS = isiOS;
    function isApple() {
        return isMacOS() || isiOS();
    }
    Util.isApple = isApple;
})(Util || (Util = {}));
var ExcelTest;
(function (ExcelTest) {
    function test_Stargarita_Run_bvt_JScript_V11() {
        var sheetName = "Sheet1";
        var arrLength = 100;
        var arrWidth = 100;
        var rangeAddress1 = "A1:cc" + arrLength;
        var valToSet = new Array(arrLength);
        for (var i = 0; i < arrLength;) {
            valToSet[i] = i;
            if (i > arrLength / 2) {
                i++;
            }
            else {
                i++;
            }
        }
        var rowHeight = 4;
        var colwidth = 4;
        var ctx = new Excel.RequestContext();
        var sheet1 = ctx.workbook.worksheets.getItem(sheetName);
        sheet1.getRange().clear();
        sheet1.getRange().format.rowHeight = 15;
        sheet1.getRange().format.columnWidth = 8;
        var format = sheet1.getCell(0, 0).getBoundingRect(sheet1.getCell(arrLength, arrWidth)).format;
        format.rowHeight = rowHeight;
        format.columnWidth = colwidth;
        format.fill.color = "black";
        ctx.load(sheet1);
        ctx.sync().then(function () {
            var chainedpromise = new OfficeExtension.Promise(function (resolve) {
                return resolve();
            });
            valToSet.map(function (item) {
                chainedpromise = chainedpromise.then(function () {
                    return Util.wait(5, function () {
                        sheet1.getCell(item - 1, item - 1).format.fill.color = "black";
                        sheet1.getCell(item, item).format.fill.color = "white";
                        ctx.sync();
                    });
                });
                return chainedpromise;
            });
        }).then(ctx.sync).then(ExcelTest.pass).catch(ExcelTest.reportError);
    }
    ExcelTest.test_Stargarita_Run_bvt_JScript_V11 = test_Stargarita_Run_bvt_JScript_V11;
    function test_Stargarita_Run2_bvt_JScript_V11() {
        var sheetName = "Sheet1";
        var arrLength = 90;
        var arrWidth = 100;
        var rangeAddress1 = "A1:cc" + arrLength;
        var valToSet = new Array(arrLength);
        for (var i = 0; i < arrLength;) {
            valToSet[i] = i;
            if (i > arrLength / 2) {
                i++;
            }
            else {
                i++;
            }
        }
        var rowHeight = 4;
        var colwidth = 4;
        var ctx = new Excel.RequestContext();
        var sheet1 = ctx.workbook.worksheets.getItem(sheetName);
        sheet1.getRange().clear();
        sheet1.getRange().format.rowHeight = 15;
        sheet1.getRange().format.columnWidth = 8;
        var format = sheet1.getCell(0, 0).getBoundingRect(sheet1.getCell(arrLength, arrWidth)).format;
        format.rowHeight = rowHeight;
        format.columnWidth = colwidth;
        format.fill.color = "black";
        ctx.load(sheet1);
        ctx.sync().then(function () {
            for (var i = 0; i < arrLength; i++) {
                Util.wait(100, function () {
                    sheet1.getCell(i - 1, i - 1).format.fill.color = "black";
                    sheet1.getCell(i, i).format.fill.color = "white";
                    ctx.sync();
                });
            }
            //var chainedpromise = new OfficeExtension.Promise(function (resolve) { return resolve(); });
            //valToSet.map(function (item) {
            //	chainedpromise = chainedpromise
            //		.then(function () {
            //			return Util.wait(5, function () {
            //				sheet1.getCell(item - 1, item - 1).format.fill.color = "black";
            //				sheet1.getCell(item, item).format.fill.color = "white";
            //				ctx.sync();
            //			});
            //		});
            //	return chainedpromise;
            //});
        }).then(ExcelTest.pass).catch(ExcelTest.reportError);
    }
    ExcelTest.test_Stargarita_Run2_bvt_JScript_V11 = test_Stargarita_Run2_bvt_JScript_V11;
    function test_Stargarita_Ru124_bvt_JScript_V11() {
        loadItems();
    }
    ExcelTest.test_Stargarita_Ru124_bvt_JScript_V11 = test_Stargarita_Ru124_bvt_JScript_V11;
    ExcelTest.gametickInMS = 100;
    ExcelTest.y_wall_top = 0;
    ExcelTest.y_wall_bottom = 98;
    ExcelTest.x_wall_left = 2;
    ExcelTest.x_wall_right = 96;
    ExcelTest.timer;
    ExcelTest.wall = {
        name: "wall",
        width: 2,
        height: 100,
        y: 0,
        x: 0,
        y_v: 0,
        x_v: 0,
        y_new: 0,
        x_new: 0,
        color: "orange",
        start_y: 0,
        start_x: 0
    };
    ExcelTest.gamegrid = {
        name: "gamegrid",
        width: 100,
        height: 100,
        y: 0,
        x: 0,
        y_v: 0,
        x_v: 0,
        y_new: 0,
        x_new: 0,
        color: "black",
        start_y: 0,
        start_x: 0
    };
    ExcelTest.ball = {
        name: "ball",
        width: 2,
        height: 2,
        y: 49,
        x: 49,
        y_v: -1,
        x_v: -2,
        y_new: 49,
        x_new: 49,
        color: "orange",
        start_y: 49,
        start_x: 49
    };
    ExcelTest.paddle = {
        name: "paddle",
        width: 2,
        height: 10,
        y: 45,
        x: 98,
        y_v: 0,
        x_v: 0,
        y_new: 45,
        x_new: 98,
        color: "orange",
        start_y: 45,
        start_x: 98
    };
    var items = [ExcelTest.gamegrid, ExcelTest.wall, ExcelTest.ball, ExcelTest.paddle];
    var redrawItems = [ExcelTest.ball, ExcelTest.paddle];
    // Update new positions with velocitx changes
    function moveItems() {
        processCollisions();
        items.forEach(function (item, indey) {
            item.y_new = item.y + item.y_v;
            item.x_new = item.x + item.x_v;
            updateWindow();
        });
    }
    ExcelTest.moveItems = moveItems;
    // Handle changes in velocitx
    function processCollisions() {
        if (ExcelTest.ball.y + ExcelTest.ball.y_v < ExcelTest.y_wall_top || ExcelTest.ball.y + ExcelTest.ball.y_v > ExcelTest.y_wall_bottom) {
            ExcelTest.ball.y_v = -ExcelTest.ball.y_v;
        }
        if (ExcelTest.ball.x + ExcelTest.ball.x_v < ExcelTest.x_wall_left) {
            ExcelTest.ball.x_v = -ExcelTest.ball.x_v;
        }
        if (ExcelTest.ball.x + ExcelTest.ball.x_v > ExcelTest.x_wall_right) {
            var ball_top_min = Math.min(ExcelTest.ball.y, ExcelTest.ball.y + ExcelTest.ball.y_v);
            var ball_bottom_max = Math.max(ExcelTest.ball.y + ExcelTest.ball.height, ExcelTest.ball.y + ExcelTest.ball.height + ExcelTest.ball.y_v);
            if ((ExcelTest.ball.y_v > 0 && ball_top_min <= ExcelTest.paddle.y && ExcelTest.paddle.y <= ball_bottom_max) || (ExcelTest.ball.y_v < 0 && ball_top_min <= ExcelTest.paddle.y + ExcelTest.paddle.height && ExcelTest.paddle.y + ExcelTest.paddle.height <= ball_bottom_max)) {
                // The ball hits a corner and bounces funnx.
                var y_v = ExcelTest.ball.y_v;
                var x_v = ExcelTest.ball.x_v;
                ExcelTest.ball.x_v = -Math.abs(y_v);
                ExcelTest.ball.y_v = -x_v;
            }
            else if (ball_bottom_max >= ExcelTest.paddle.y && ball_top_min <= ExcelTest.paddle.y + ExcelTest.paddle.height) {
                ExcelTest.ball.x_v = -ExcelTest.ball.x_v;
            }
            else {
                clearInterval(ExcelTest.timer);
                RichApiTest.log.comment("You lost");
            }
        }
    }
    ExcelTest.processCollisions = processCollisions;
    // Draw things!
    function updateWindow() {
        var sheetName = "Sheet1";
        var ctx = new Excel.RequestContext();
        var sheet1 = ctx.workbook.worksheets.getItem(sheetName);
        redrawItems.forEach(function (item, indey) {
            sheet1.getCell(item.y, item.x).getBoundingRect(sheet1.getCell(item.y + item.height - 1, item.x + item.width - 1)).format.fill.color = ExcelTest.gamegrid.color;
            sheet1.getCell(item.y_new, item.x_new).getBoundingRect(sheet1.getCell(item.y_new + item.height - 1, item.x_new + item.width - 1)).format.fill.color = item.color;
            ctx.sync();
            //$("#" + item.name).css({
            //	"top": item.y_new,
            //	"bottom": item.x_new
            //})
            item.y = item.y_new;
            item.x = item.x_new;
        });
        $("#debug").html("y: " + ExcelTest.ball.y + "<br/>y_v: " + ExcelTest.ball.y_v + "<br/>x: " + ExcelTest.ball.x + "<br/>x_v: " + ExcelTest.ball.x_v);
    }
    ExcelTest.updateWindow = updateWindow;
    // Initialize the board
    function loadItems() {
        var sheetName = "Sheet1";
        var ctx = new Excel.RequestContext();
        var sheet1 = ctx.workbook.worksheets.getItem(sheetName);
        items.forEach(function (item, indey) {
            item.y = item.start_y;
            item.x = item.start_x;
            sheet1.getCell(item.y, item.x).getBoundingRect(sheet1.getCell(item.y + item.height - 1, item.x + item.width - 1)).format.fill.color = item.color;
            ctx.sync();
            //jQuery('<div></div>', {
            //	id: item.name
            //})
            //	.width(item.width)
            //	.height(item.height)
            //	.css({
            //		"position": "absolute",
            //		"top": item.y + "px",
            //		"bottom": item.x + "px",
            //		"background-color": item.color
            //	})
            //	.appendTo("#game");
        });
        // Add a kexpress handler. 
        $('html').keydown(function (e) {
            var event = window.event ? window.event : e;
            switch (event.keyCode) {
                case 38:
                    if (ExcelTest.paddle.y > ExcelTest.y_wall_top) {
                        ExcelTest.paddle.y_new = ExcelTest.paddle.y - 1;
                        updateWindow();
                    }
                    break;
                case 40:
                    if (ExcelTest.paddle.y + ExcelTest.paddle.height < ExcelTest.y_wall_bottom) {
                        ExcelTest.paddle.y_new = ExcelTest.paddle.y + 1;
                        updateWindow();
                    }
                    break;
            }
        });
        ExcelTest.timer = setInterval(moveItems, 100);
    }
    ExcelTest.loadItems = loadItems;
})(ExcelTest || (ExcelTest = {}));

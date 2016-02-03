var RichApiTest;
(function (RichApiTest) {
    var RestUtility = (function () {
        function RestUtility() {
        }
        RestUtility.getBaseUrlUsingOverride = function (defaultBaseUrl) {
            var elemBaseUrl = document.getElementById("TxtRichApiRestBaseUrlOverride");
            var url = elemBaseUrl.value;
            url = OfficeExtension.Utility.trim(url);
            if (url.length == 0) {
                url = defaultBaseUrl;
            }
            if (url.charAt(url.length - 1) !== "/") {
                url = url + "/";
            }
            return url;
        };
        RestUtility.updateRequestInfoUsingProxyIfNecessary = function (request) {
            var method = request.method;
            if (!method) {
                method = "GET";
            }
            method = method.toUpperCase();
            if (!request.headers) {
                request.headers = {};
            }
            if (OfficeExtension.Utility.trim(jQuery("#TxtRichApiHeaderName1").val()).length > 0) {
                request.headers[OfficeExtension.Utility.trim(jQuery("#TxtRichApiHeaderName1").val())] = OfficeExtension.Utility.trim(jQuery("#TxtRichApiHeaderValue1").val());
            }
            if (OfficeExtension.Utility.trim(jQuery("#TxtRichApiHeaderName2").val()).length > 0) {
                request.headers[OfficeExtension.Utility.trim(jQuery("#TxtRichApiHeaderName2").val())] = OfficeExtension.Utility.trim(jQuery("#TxtRichApiHeaderValue2").val());
            }
            if (OfficeExtension.Utility.trim(jQuery("#TxtRichApiHeaderName3").val()).length > 0) {
                request.headers[OfficeExtension.Utility.trim(jQuery("#TxtRichApiHeaderName3").val())] = OfficeExtension.Utility.trim(jQuery("#TxtRichApiHeaderValue3").val());
            }
            var url = request.url;
            if (url.indexOf(OfficeExtension.Constants.localDocumentApiPrefix) < 0) {
                var proxyUrl = "/RichApiRestProxy.ashx?RequestUrl=" + encodeURIComponent(url);
                if (document.getElementById("ChkRichApiProxyUseFiddler").checked) {
                    proxyUrl = proxyUrl + "&UseFiddler=1";
                }
                url = proxyUrl;
                if (method === "PATCH" || method === "DELETE") {
                    request.headers["X-HTTP-METHOD"] = method;
                    method = "POST";
                }
            }
            request.url = url;
            request.method = method;
        };
        RestUtility.invoke = function (request) {
            RestUtility.updateRequestInfoUsingProxyIfNecessary(request);
            var option = {};
            option.url = request.url;
            option.type = request.method;
            option.cache = false;
            if (request.headers && request.headers["CONTENT-TYPE"]) {
                option.contentType = request.headers["CONTENT-TYPE"];
                delete request.headers["CONTENT-TYPE"];
            }
            else {
                option.contentType = "application/json";
            }
            if (request.method === "POST" || request.method == "PATCH" || request.method == "PUT") {
                option.data = request.body;
            }
            option.headers = request.headers;
            var ret = new OfficeExtension['Promise'](function (resolve, reject) {
                jQuery.ajax(option).then(function (data, textStatus, jqXHR) {
                    var resp = {
                        statusCode: jqXHR.status,
                        body: jqXHR.responseText,
                        headers: {}
                    };
                    resp.headers = RestUtility.parseHeaders(jqXHR.getAllResponseHeaders());
                    RestUtility.logRestResponse(resp);
                    resolve(resp);
                }, function (jqXHR, textStatus, errorThrown) {
                    var resp = {
                        statusCode: jqXHR.status,
                        body: jqXHR.responseText,
                        headers: {}
                    };
                    resp.headers = RestUtility.parseHeaders(jqXHR.getAllResponseHeaders());
                    RestUtility.logRestResponse(resp);
                    resolve(resp);
                });
            });
            return ret;
        };
        RestUtility.verifyStatusCodeThrow = function (respInfo, expectedCode) {
            if (!RestUtility.verifyStatusCode(respInfo, expectedCode)) {
                throw Error("Expected status code " + expectedCode + ", but received " + respInfo.statusCode);
            }
        };
        RestUtility.verifyStatusCode = function (respInfo, expectedCode) {
            var success = false;
            if (expectedCode) {
                success = (expectedCode == respInfo.statusCode);
                if (!success) {
                    RichApiTest.log.comment("Expected status code " + expectedCode + ", but received " + respInfo.statusCode);
                }
            }
            else {
                if (respInfo.statusCode == RestUtility.httpStatusOK || respInfo.statusCode == RestUtility.httpStatusCreated || respInfo.statusCode == RestUtility.httpStatusNoContent) {
                    success = true;
                }
                else {
                    RichApiTest.log.comment("Failed request");
                    success = false;
                }
            }
            return success;
        };
        RestUtility.logRestResponse = function (respInfo) {
            RichApiTest.log.comment("Status:" + respInfo.statusCode);
            RichApiTest.log.comment("Body:" + respInfo.body);
        };
        RestUtility.parseHeaders = function (allResponseHeaders) {
            var ret = {};
            var regex = new RegExp("\r?\n");
            var entries = allResponseHeaders.split(regex);
            for (var i = 0; i < entries.length; i++) {
                var entry = entries[i];
                if (typeof (entry) === "string" && entry.length > 0) {
                    var index = entry.indexOf(':');
                    if (index > 0) {
                        var key = entry.substr(0, index);
                        var value = entry.substr(index + 1);
                        key = OfficeExtension.Utility.trim(key);
                        value = OfficeExtension.Utility.trim(value);
                        ret[key.toUpperCase()] = value;
                    }
                }
            }
            return ret;
        };
        RestUtility.httpStatusOK = 200;
        RestUtility.httpStatusCreated = 201;
        RestUtility.httpStatusNoContent = 204;
        RestUtility.httpStatusBadRequest = 400;
        RestUtility.httpStatusNotFound = 404;
        RestUtility.httpStatusMethodNotAllowed = 405;
        RestUtility.httpMethodGet = "GET";
        RestUtility.httpMethodPost = "POST";
        RestUtility.httpMethodPatch = "PATCH";
        RestUtility.httpMethodDelete = "DELETE";
        return RestUtility;
    })();
    RichApiTest.RestUtility = RestUtility;
    var Settings = (function () {
        function Settings() {
            this.timeoutSeconds = 60;
        }
        return Settings;
    })();
    RichApiTest.Settings = Settings;
    var Logger = (function () {
        function Logger() {
        }
        Logger.prototype.comment = function (message) {
            var elem = document.createElement("div");
            elem.innerText = message;
            document.getElementById("DivRichApiTestResult").appendChild(elem);
        };
        Logger.prototype.clear = function () {
            document.getElementById("DivRichApiTestResult").innerHTML = "";
            this.clearDone();
        };
        Logger.prototype.clearDone = function () {
            document.getElementById("TxtRichApiDone").value = "";
        };
        Logger.prototype.done = function (success) {
            if (success) {
                this.comment("Success");
                document.getElementById("TxtRichApiDone").value = "S";
            }
            else {
                this.comment("Fail");
                document.getElementById("TxtRichApiDone").value = "F";
            }
            if (this.currentTestName != null) {
                findTestButtonAndSetItsColor(this.currentTestName, (success ? "green" : "red"));
            }
        };
        Logger.prototype.fail = function (message) {
            this.comment(message);
            this.done(false);
        };
        Logger.prototype.pass = function (message) {
            this.comment(message);
            this.done(true);
        };
        return Logger;
    })();
    RichApiTest.Logger = Logger;
    RichApiTest.log = new Logger();
    RichApiTest.settings = new Settings();
    var UIConstants;
    (function (UIConstants) {
        UIConstants.TxtRichApiAgsUrl = "TxtRichApiAgsUrl";
        UIConstants.TxtRichApiAgsFileName = "TxtRichApiAgsFileName";
        UIConstants.TxtRichApiHeaderName1 = "TxtRichApiHeaderName1";
        UIConstants.TxtRichApiHeaderName2 = "TxtRichApiHeaderName2";
        UIConstants.TxtRichApiHeaderName3 = "TxtRichApiHeaderName3";
        UIConstants.TxtRichApiHeaderValue1 = "TxtRichApiHeaderValue1";
        UIConstants.TxtRichApiHeaderValue2 = "TxtRichApiHeaderValue2";
        UIConstants.TxtRichApiHeaderValue3 = "TxtRichApiHeaderValue3";
        UIConstants.TxtRichApiRestBaseUrlOverride = "TxtRichApiRestBaseUrlOverride";
    })(UIConstants || (UIConstants = {}));
    function buildUI(parent, testNs, intellisensePaths) {
        if (intellisensePaths === void 0) { intellisensePaths = []; }
        var div = document.createElement("div");
        div.innerHTML = "" + "<div id='script-or-rest-resizer' class='ui-widget-content' style='padding: 5px; width: calc(100%-20px); height:50px;'>" + "	<div id='script-or-rest-accordion'>" + "		<h3>" + "			Script Editor:" + "			&nbsp;&nbsp;&nbsp;" + "			<button onclick='RichApiTest.invokeScriptEditorRunButton()'>Run</button>" + "			<button onclick='RichApiTest.setNewScript(\"" + testNs + "\")' style='margin-left: 15px;'>New</button>" + "			<button onclick='javascript:window.location.reload()' style='margin-left: 15px;'>Reload app</button>" + "		</h3>" + "		<div id='TxtRichApiScript' style='padding:0'></div>" + "	</div>" + "</div>" + "<br />" + "<div>" + "	<b>Result:</b>&nbsp;&nbsp;<input type='text' id='TxtRichApiDone' size='2' / > " + "	<button onclick = 'RichApiTest.log.clear()' > Clear Log </button >" + "	&nbsp;&nbsp;<input id='ChkRichApiBreakOnFailure' class='simple-save-restore' type='checkbox' /><label for='ChkRichApiBreakOnFailure'>Break on failure</label>" + "</div>" + "<div id='DivRichApiTestResult'></div>" + "<hr />" + "<div id='DivRichApiTestButtonSelection'>" + "<br /><select id='InpRichApiGroups' class='simple-save-restore' onChange='RichApiTest && RichApiTest.toggleCurrentGroupSelection && RichApiTest.toggleCurrentGroupSelection()'></select>" + "&nbsp;<input type='Checkbox' name='ChkJScript' checked='checked' id='ChkJScript' class='simple-save-restore' onclick='RichApiTest.toggleCheckboxFilters();' /><label for='ChkJScript'>JavaScript</label>" + "&nbsp;<input type='Checkbox' name='ChkREST' checked='checked' id='ChkREST' class='simple-save-restore' onclick='RichApiTest.toggleCheckboxFilters();' /><label for='ChkREST'>REST</label>" + "<div id='DivRichApiVersionsSelection'>Version: </div>" + "</div>" + "<hr />" + "<div id='DivRichApiTests'></div>";
        $(parent).empty();
        parent.appendChild(div);
        var savedApiGroupName = getLocalSettingIfAny("InpRichApiGroups");
        var savedVersionChkboxes = getLocalStorageSettingForVersionChkboxes();
        restoreLocalSetting('TxtRichApiRestBaseUrlOverride', function (value) {
            $("#TxtRichApiRestBaseUrlOverride").val(value || "http://document.localhost/_api/");
        });
        $('#TxtRichApiRestBaseUrlOverride').blur(function () {
            if (!($(this).val())) {
                $(this).val("http://document.localhost/_api/");
            }
            storeLocalSetting("TxtRichApiRestBaseUrlOverride", $(this).val());
        });
        $(function () {
            var $accordionElem = $("#script-or-rest-accordion");
            $accordionElem.accordion({
                heightStyle: "fill",
                animate: 100,
                activate: function () {
                    MonacoEditorIntegration.resizeEditor(true);
                    storeLocalSetting('script-or-rest-accordion', $accordionElem.accordion("option", "active"));
                }
            });
            restoreLocalSetting('script-or-rest-accordion', function (value) {
                if (value) {
                    $accordionElem.accordion("option", "active", parseInt(value));
                }
            });
            $("#script-or-rest-resizer").resizable({
                handles: "s",
                resize: function () {
                    $accordionElem.accordion("refresh");
                    MonacoEditorIntegration.resizeEditor();
                }
            });
            MonacoEditorIntegration.initializeJsEditor('TxtRichApiScript', intellisensePaths);
        });
        $("#DivRichApiAllTestResult, #DivRichApiTestResult").dblclick(function () {
            var elementId = this.id;
            if (document.body.createTextRange) {
                var docRange = document.body.createTextRange();
                docRange.moveToElementText(document.getElementById(elementId));
                docRange.select();
            }
            else if (window.getSelection) {
                var windowRange = document.createRange();
                windowRange.selectNode(document.getElementById(elementId));
                window.getSelection().addRange(windowRange);
            }
        });
        appendTests(testNs);
        $('.simple-save-restore').each(function () {
            addSimpleSaveRestore($(this));
        });
        if (savedApiGroupName) {
            $("#InpRichApiGroups").val(savedApiGroupName);
        }
        if (savedVersionChkboxes) {
            for (var key in savedVersionChkboxes) {
                $("#" + key).val(savedVersionChkboxes[key]);
            }
        }
        toggleCurrentGroupSelection();
        toggleCheckboxFilters();
    }
    RichApiTest.buildUI = buildUI;
    function addSimpleSaveRestore($elem) {
        var id = $elem.attr("id");
        if (!id) {
            throw new Error('Saveable/restorable element must have an ID.');
        }
        var tagName = $elem[0].tagName.toLowerCase();
        var inputType = (tagName == "input" && $elem.attr("type")) ? $elem.attr("type").toLowerCase() : undefined;
        var isAcceptableElement = (tagName == "select" || tagName == "textarea" || (tagName == "input" && (inputType === undefined || inputType == "" || inputType == "text" || inputType == "checkbox")));
        if (!isAcceptableElement) {
            throw new Error('Invalid element type for saving/restoring. Element with id ' + id);
        }
        restoreLocalSetting(id, function (value) {
            if (value != null) {
                setValue(value);
            }
        });
        $elem.blur(function () {
            storeLocalSetting(id, getValue());
        });
        function setValue(value) {
            if (tagName == "input" && inputType == "checkbox") {
                document.getElementById(id).checked = ((value === undefined || value == "true") ? true : false);
            }
            else {
                $elem.val(value);
            }
        }
        function getValue() {
            if (tagName == "input" && inputType == "checkbox") {
                return document.getElementById(id).checked;
            }
            else {
                return $elem.val();
            }
        }
    }
    RichApiTest.groups = [];
    RichApiTest.testVersions = [];
    function appendTests(ns) {
        var parts = ns.split(".");
        var obj = window;
        for (var i = 0; i < parts.length; i++) {
            obj = obj[parts[i]];
        }
        var div = document.getElementById("DivRichApiTests");
        var anyFound = false;
        for (var f in obj) {
            if (typeof (f) == "string" && typeof (obj[f]) == "function") {
                var fn = obj[f];
                var name = f.toString();
                if (name.indexOf("test") == 0) {
                    var tokens = name.split("_");
                    var groupName = "";
                    var buttonName = name;
                    var typeName = "JScript";
                    var testVersion = "VUnknown";
                    if (tokens.length > 1) {
                        groupName = tokens[1];
                    }
                    if (tokens.length > 2) {
                        buttonName = tokens[1] + "_" + tokens[2];
                    }
                    if (tokens.length > 4) {
                        typeName = tokens[4];
                        if (typeName != "JScript" && typeName != "REST") {
                            typeName = "JScript";
                        }
                        buttonName += "_" + typeName.charAt(0);
                    }
                    if (tokens.length > 5) {
                        testVersion = tokens[5];
                    }
                    if (arrayPushUnique(RichApiTest.groups, groupName)) {
                        var newgroupdiv = document.createElement("div");
                        newgroupdiv.id = "DivRichApiTests_" + groupName;
                        div.appendChild(newgroupdiv);
                    }
                    arrayPushUnique(RichApiTest.testVersions, testVersion);
                    var groupdiv = document.getElementById("DivRichApiTests_" + groupName);
                    anyFound = true;
                    var button = document.createElement("button");
                    button.innerHTML = buttonName;
                    var fullName = ns + "." + name;
                    button.setAttribute("name", name);
                    button.setAttribute("fullname", fullName);
                    button.setAttribute("testtype", typeName);
                    button.setAttribute("testversion", testVersion);
                    button.onclick = new Function("RichApiTest.invokeOneTest('" + name + "', '" + fullName + "');");
                    groupdiv.appendChild(button);
                }
            }
        }
        if (!anyFound) {
            $(div).html("Could not find any tests in the " + ns + " namespace." + "<br/><br/>" + "Please be sure that you've built the JavaScript tests (e.g., perform an \"obuild\" on " + "\"%SRCROOT%\\xlshared\\src\\Api\\test\\JScript\\\"), " + "<br/>then run<br/>" + "\"%SRCROOT%\\osfclient\\RichApi\\Test\\RichApiAgaveWeb\\CopyJsFromTarget.bat\"" + "<br/> and try again.");
        }
        else {
            var inpSel = document.getElementById("InpRichApiGroups");
            for (var i = 0; i < RichApiTest.groups.length; i++) {
                var opt = RichApiTest.groups[i];
                var el = document.createElement("option");
                el.textContent = opt.length == 0 ? "(none)" : opt;
                el.value = opt;
                inpSel.appendChild(el);
            }
            var divSelection = document.getElementById("DivRichApiVersionsSelection");
            for (var i = 0; i < RichApiTest.testVersions.length; i++) {
                var testVersionName = RichApiTest.testVersions[i];
                var inpElement = document.createElement("input");
                inpElement.type = "Checkbox";
                inpElement.name = "ChkVersion_" + testVersionName;
                inpElement.checked = true;
                inpElement.id = "ChkVersion_" + testVersionName;
                inpElement.className = "simple-save-restore";
                inpElement.onclick = function () {
                    RichApiTest.toggleCheckboxFilters();
                };
                var lblElement = document.createElement("label");
                lblElement.htmlFor = "ChkVersion_" + testVersionName;
                var labelName = testVersionName.substr(1, 1) + "." + testVersionName.substr(2);
                lblElement.innerText = labelName;
                divSelection.appendChild(inpElement);
                divSelection.appendChild(lblElement);
            }
            hideAllTestGroups();
            toggleCurrentGroupSelection();
        }
    }
    RichApiTest.appendTests = appendTests;
    function arrayPushUnique(arr, item) {
        if (arr.indexOf(item) == -1) {
            arr.push(item);
            return true;
        }
        return false;
    }
    RichApiTest.arrayPushUnique = arrayPushUnique;
    function arrayPushUniqueWithKey(arr, item, key) {
        if (arr.indexOf(item) == -1 && arr[key] == undefined) {
            arr[key] = item;
            return true;
        }
        return false;
    }
    RichApiTest.arrayPushUniqueWithKey = arrayPushUniqueWithKey;
    function toggleCurrentGroupSelection() {
        hideAllTestGroups();
        var inpSelVal = document.getElementById("InpRichApiGroups").value;
        var div = document.getElementById("DivRichApiTests_" + inpSelVal);
        var state = document.getElementById("DivRichApiTests_" + inpSelVal).style.display;
        if (state == "none") {
            document.getElementById("DivRichApiTests_" + inpSelVal).style.display = "block";
        }
        else {
            document.getElementById("DivRichApiTests_" + inpSelVal).style.display = "none";
        }
    }
    RichApiTest.toggleCurrentGroupSelection = toggleCurrentGroupSelection;
    function toggleCheckboxFilters() {
        var elements = document.getElementById("DivRichApiTests").getElementsByTagName("button");
        var names = [];
        var testTypesToHide = new Array();
        if (!document.getElementById("ChkJScript").checked) {
            arrayPushUnique(testTypesToHide, "JScript");
        }
        if (!document.getElementById("ChkREST").checked) {
            arrayPushUnique(testTypesToHide, "REST");
        }
        var testVersionsToHide = new Array();
        var versionChkboxes = document.getElementById("DivRichApiVersionsSelection").getElementsByTagName("input");
        for (var i = 0; i < versionChkboxes.length; i++) {
            if (!versionChkboxes[i].checked) {
                var cbVersion = versionChkboxes[i].id.substr(versionChkboxes[i].id.indexOf("_") + 1);
                arrayPushUnique(testVersionsToHide, cbVersion);
            }
        }
        for (var i = 0; i < elements.length; i++) {
            var testtype = elements[i].getAttribute("testtype");
            var testversion = elements[i].getAttribute("testversion");
            if ($.inArray(testversion, testVersionsToHide) != -1 || $.inArray(testtype, testTypesToHide) != -1) {
                elements[i].style.display = "none";
            }
            else {
                elements[i].style.display = "inline";
            }
        }
    }
    RichApiTest.toggleCheckboxFilters = toggleCheckboxFilters;
    function hideAllTestGroups() {
        for (var i = 0; i < RichApiTest.groups.length; i++) {
            document.getElementById("DivRichApiTests_" + RichApiTest.groups[i]).style.display = "none";
        }
    }
    RichApiTest.hideAllTestGroups = hideAllTestGroups;
    function showAllTestGroups() {
        for (var i = 0; i < RichApiTest.groups.length; i++) {
            document.getElementById("DivRichApiTests_" + RichApiTest.groups[i]).style.display = "block";
        }
    }
    RichApiTest.showAllTestGroups = showAllTestGroups;
    function toggleAllTestButtons() {
        RichApiTest.log.clear();
        for (var i = 0; i < RichApiTest.groups.length; i++) {
            var state = document.getElementById("DivRichApiTests_" + RichApiTest.groups[i]).style.display;
            if (state == "none") {
                RichApiTest.log.comment("Changing visibility of " + document.getElementById("DivRichApiTests_" + RichApiTest.groups[i]).id + " to visible");
                document.getElementById("DivRichApiTests_" + RichApiTest.groups[i]).style.display = "block";
            }
            else {
                RichApiTest.log.comment("Changing visibility of " + document.getElementById("DivRichApiTests_" + RichApiTest.groups[i]).id + " to hidden");
                document.getElementById("DivRichApiTests_" + RichApiTest.groups[i]).style.display = "none";
            }
        }
    }
    RichApiTest.toggleAllTestButtons = toggleAllTestButtons;
    function invokeScriptEditorRunButton() {
        RichApiTest.log.clear();
        RichApiTest.log.currentTestName = null;
        RichApiTest.log.comment("--Script--");
        var editorText = MonacoEditorIntegration.getEditorValue();
        if (isTrulyJavaScript(editorText)) {
            evaluateJavaScriptCommon(createJavaScriptEvalAction(editorText));
        }
        else {
            MonacoEditorIntegration.getEditorTextAsJavaScript().then(function (output) {
                if (output == null) {
                    RichApiTest.log.comment("Invalid JavaScript / TypeScript. Please fix the errors shown in the code editor and try again.");
                }
                else {
                    evaluateJavaScriptCommon(createJavaScriptEvalAction(output.content));
                }
            });
        }
        function isTrulyJavaScript(text) {
            try {
                new Function(text);
                return true;
            }
            catch (syntaxError) {
                return false;
            }
        }
    }
    RichApiTest.invokeScriptEditorRunButton = invokeScriptEditorRunButton;
    function createJavaScriptEvalAction(script) {
        var FakeConsole = (function () {
            function FakeConsole() {
            }
            FakeConsole.prototype.log = function (text) {
                RichApiTest.log.comment(text);
            };
            FakeConsole.prototype.error = function (text) {
                RichApiTest.log.comment(text);
            };
            return FakeConsole;
        })();
        var evalAction = (function (script, console) {
            return function () {
                eval(script);
            };
        })(script, new FakeConsole());
        return evalAction;
    }
    function evaluateJavaScriptCommon(evaluationAction) {
        try {
            evaluationAction();
        }
        catch (ex) {
            RichApiTest.log.comment("Exception thrown during code execution:");
            RichApiTest.log.comment(ex.toString());
            if (ex.stack) {
                RichApiTest.log.comment(ex.stack);
            }
        }
    }
    function setNewScript(testNs) {
        if (window[testNs] && window[testNs]["getNewScriptText"]) {
            MonacoEditorIntegration.setEditorValue(window[testNs]["getNewScriptText"]());
        }
        else {
            var errorText = "RichApiTest.log.comment('No \"getNewScriptText()\" function is specified in the \"" + testNs + "\" namespace. Please define it in your test suite (e.g., in \"Common.ts\"), first.');";
            MonacoEditorIntegration.setEditorValue(errorText);
        }
    }
    RichApiTest.setNewScript = setNewScript;
    function invokeRest() {
        var elemBaseUrl = document.getElementById("TxtRichApiRestBaseUrlOverride");
        var elemPath = document.getElementById("TxtRichApiRestPath");
        var url = elemBaseUrl.value;
        url = OfficeExtension.Utility.trim(url);
        if (url.length == 0) {
            RichApiTest.log.comment("Missed baseUrl");
            return;
        }
        if (url.charAt(url.length - 1) !== "/") {
            url = url + "/";
        }
        var path = OfficeExtension.Utility.trim(elemPath.value);
        if (path.length > 0 && path.charAt(0) === "/") {
            path = path.substr(1);
        }
        url = url + path;
        var requestInfo = {
            url: url,
            method: jQuery("#DdlRichApiRestMethod").val(),
            body: jQuery("#TxtRichApiRestBody").val(),
            headers: {}
        };
        RestUtility.invoke(requestInfo);
    }
    RichApiTest.invokeRest = invokeRest;
    function invokeOneTest(testName, funcName) {
        var fn = eval(funcName);
        RichApiTest.log.clear();
        RichApiTest.log.currentTestName = funcName;
        findTestButtonAndSetItsColor(funcName, "orange");
        RichApiTest.log.comment("--" + testName + "--");
        var funcEntire = fn.toString();
        var funcBody = funcEntire.substring(funcEntire.indexOf("{") + 1, funcEntire.lastIndexOf("}"));
        funcBody = replaceAll(funcBody, ".then(", "\r\n.then(");
        funcBody = replaceAll(funcBody, ".catch(", "\r\n.catch(");
        var lines = funcBody.split('\n');
        var isFirstNonEmptyLineFound = false;
        var firstNonEmptyLineIndex = 0;
        var minimumIndexOfNonWhitespace = Number.MAX_VALUE;
        var nonWhiteSpacePattern = /[\S]/;
        for (var i = 0; i < lines.length; i++) {
            var nonWhitespaceIndexWithinLine = lines[i].search(nonWhiteSpacePattern);
            if (nonWhitespaceIndexWithinLine >= 0) {
                isFirstNonEmptyLineFound = true;
                minimumIndexOfNonWhitespace = Math.min(minimumIndexOfNonWhitespace, nonWhitespaceIndexWithinLine);
            }
            else {
                if (!isFirstNonEmptyLineFound) {
                    firstNonEmptyLineIndex = i;
                }
            }
        }
        lines = lines.slice(firstNonEmptyLineIndex + 1);
        funcBody = lines.map(function (line) {
            return line.substr(minimumIndexOfNonWhitespace);
        }).join('\n');
        MonacoEditorIntegration.setEditorValue(funcBody);
        if (funcBody.indexOf("console.log") == -1) {
            evaluateJavaScriptCommon(function () {
                fn();
            });
        }
        else {
            RichApiTest.log.comment('NOTE: function contains "console.log" calls, so evaluating using "eval". Breakpoints won\'t work in this mode. ' + 'To pause execution, add a "debugger;" statement to the editor, instead.');
            evaluateJavaScriptCommon(createJavaScriptEvalAction(funcBody));
        }
    }
    RichApiTest.invokeOneTest = invokeOneTest;
    function clearAllTestResult() {
        document.getElementById("DivRichApiAllTestResult").innerHTML = "";
        var buttons = document.getElementById("DivRichApiTests").getElementsByTagName("button");
        for (var i = 0; i < buttons.length; i++) {
            $(buttons[i]).css("color", "black");
        }
    }
    RichApiTest.clearAllTestResult = clearAllTestResult;
    function invokeAllTests() {
        invokeTests(function (name) {
            return true;
        });
    }
    RichApiTest.invokeAllTests = invokeAllTests;
    function invokeTestsWithKeywords() {
        var keywordsText = document.getElementById("TxtRichApiTestKeywords").value;
        var keywords = keywordsText.split(/[\s,]+/);
        for (var i = 0; i < keywords.length; i++) {
            keywords[i] = keywords[i].trim().toLowerCase();
        }
        invokeTests(function (name) {
            var ret = true;
            name = name.toLowerCase();
            var charCodeNot = "!".charCodeAt(0);
            for (var i = 0; i < keywords.length; i++) {
                var keyword = keywords[i];
                if (keyword.length > 0) {
                    if (keyword.length > 1 && keyword.charCodeAt(0) == charCodeNot) {
                        if (name.indexOf(keyword.substr(1)) >= 0) {
                            ret = false;
                            break;
                        }
                    }
                    else {
                        if (name.indexOf(keyword) < 0) {
                            ret = false;
                            break;
                        }
                    }
                }
            }
            return ret;
        });
    }
    RichApiTest.invokeTestsWithKeywords = invokeTestsWithKeywords;
    function invokeTests(selectFun) {
        var elements = document.getElementById("DivRichApiTests").getElementsByTagName("button");
        var names = [];
        var fullNames = [];
        for (var i = 0; i < elements.length; i++) {
            var name = elements[i].getAttribute("name");
            if (selectFun(name)) {
                names.push(name);
                fullNames.push(elements[i].getAttribute("fullname"));
            }
        }
        if (names.length == 0) {
            RichApiTest.log.comment("---No Tests to run---");
            return;
        }
        RichApiTest.log.clearDone();
        clearAllTestResult();
        var count = names.length;
        var testPassCount = 0;
        var testFailCount = 0;
        var testTimeoutCount = 0;
        var timeoutSeconds = RichApiTest.settings.timeoutSeconds;
        var intervalMilliseconds = 100;
        var timeoutCount = timeoutSeconds * 1000 / intervalMilliseconds;
        var timeoutIndex = 0;
        var indexCurrentTest = 0;
        invokeOneTest(names[indexCurrentTest], fullNames[indexCurrentTest]);
        var handle = window.setInterval(function () {
            var doneElem = document.getElementById("TxtRichApiDone");
            var couldStart = false;
            if (doneElem.value == 'S') {
                var elem = document.createElement("div");
                elem.innerText = "Succeeded " + names[indexCurrentTest];
                $(elem).css("color", "green");
                document.getElementById("DivRichApiAllTestResult").appendChild(elem);
                couldStart = true;
                testPassCount++;
            }
            else if (doneElem.value == 'F') {
                var elem = document.createElement("div");
                elem.innerText = "Failed " + names[indexCurrentTest];
                $(elem).css("color", "red");
                document.getElementById("DivRichApiAllTestResult").appendChild(elem);
                copyResultToAllTestResult();
                couldStart = true;
                testFailCount++;
            }
            else if (timeoutIndex > timeoutCount) {
                var elem = document.createElement("div");
                elem.innerText = "Timeout " + names[indexCurrentTest];
                $(elem).css("color", "orange");
                document.getElementById("DivRichApiAllTestResult").appendChild(elem);
                copyResultToAllTestResult();
                couldStart = true;
                testTimeoutCount++;
            }
            if (couldStart) {
                indexCurrentTest++;
                var shouldContinue = true;
                if ((testFailCount > 0 || testTimeoutCount > 0) && document.getElementById('ChkRichApiBreakOnFailure').checked) {
                    shouldContinue = false;
                }
                if (shouldContinue && indexCurrentTest < count) {
                    invokeOneTest(names[indexCurrentTest], fullNames[indexCurrentTest]);
                    timeoutIndex = 0;
                }
                else {
                    window.clearInterval(handle);
                    RichApiTest.log.comment("---Done all tests---");
                    var elem = document.createElement("div");
                    elem.innerText = "---Done all tests--- (" + count + " total; " + testPassCount + " passed, " + testFailCount + " failed, " + testTimeoutCount + " timed out)";
                    document.getElementById("DivRichApiAllTestResult").appendChild(elem);
                }
            }
            timeoutIndex++;
        }, 100);
    }
    function copyResultToAllTestResult() {
        var elem = document.createElement("div");
        elem.innerHTML = document.getElementById("DivRichApiTestResult").innerHTML;
        document.getElementById("DivRichApiAllTestResult").appendChild(elem);
    }
    function findTestButtonAndSetItsColor(testFullName, color) {
        var buttons = document.getElementById("DivRichApiTests").getElementsByTagName("button");
        for (var i = 0; i < buttons.length; i++) {
            if (buttons[i].getAttribute("fullName") == testFullName) {
                $(buttons[i]).css("color", color);
                return;
            }
        }
    }
    function replaceAll(text, search, replace) {
        return text.split(search).join(replace);
    }
    function storeLocalSetting(key, value) {
        if (window.localStorage) {
            window.localStorage[key] = value;
        }
    }
    function getLocalSettingIfAny(key) {
        if (window.localStorage) {
            return window.localStorage[key];
        }
        else {
            return undefined;
        }
    }
    function getLocalStorageSettingForVersionChkboxes() {
        if (window.localStorage) {
            var versionChkboxes = new Array();
            for (var key in window.localStorage) {
                if (key.indexOf("ChkVersion_") != -1) {
                    arrayPushUniqueWithKey(versionChkboxes, window.localStorage[key], key);
                }
            }
            return versionChkboxes;
        }
        else {
            return undefined;
        }
    }
    function restoreLocalSetting(key, action) {
        action(getLocalSettingIfAny(key));
    }
    function pasteOAuthToken() {
        if (window.clipboardData) {
            var content = window.clipboardData.getData('Text');
            jQuery("#TxtRichApiHeaderName1").val("Authorization").trigger("blur");
            jQuery("#TxtRichApiHeaderValue1").val("Bearer " + content).trigger("blur");
        }
    }
    RichApiTest.pasteOAuthToken = pasteOAuthToken;
    function initGraphSettings() {
        var tokenServiceUrl = "";
        var clientId = "";
        var refreshToken = "";
        var workloadRefreshToken = "";
        var authEnv = jQuery("#DdlAuthEnv").val();
        if (authEnv == "ppe" || authEnv == "dev") {
            tokenServiceUrl = "https://login.windows-ppe.net/common/oauth2/token";
            clientId = "09d9cc54-6048-4c79-b468-99aa29c6e98d";
            refreshToken = "AAABAAAAo3ZCPl0FaU2WWRdLWLHperA8sJ4PqXDxCTLjPNRJsutVXPEEEc-q4h3YgZ2IUx9ogcH0iUE7juPkQGt_9kW7UIKmhfoye0ob3Y629xtAFc20jv3mO1cSQlKzuaPjjwIg91RQ1MbKbBqVLKeWRJ62MYJoBH4pnsLQXbv_H4hpENnIfT4CKSbDA4MCKhjXzL1TyCBSAFfjU-5ddUvyj_m2HkIL0mdysjkDpLY4cMNr1gBVxW4isHYkR23pGZsVJdVgJgCJ_k4Gf49Pypzlor6qSynu3w9TtlEZsKswMLFqKKNqnMYJh6eSLh7Q3ljXW21iDmsxXaT-BTiuBwrJN4if3oRHyVbo4IeNHzc3dHrsBjlfkR8LdhrdPvoZz9OD7RYaopaN-mAtZplN16I-pev_ii6Y73FCPp3yKDXNoIhJC2O-Wcgl8Ev0CPOeSq8tdtfE-VE53SIgZnc0MjE4WiZzFyejzatXDIhI9XQAXJC5JPGhL1q6AYtoP4Zih_sLDywxitrU9XikneZyjy1RGmmxMzuOjyafXZnlTLLD7ko7XYADZNps7J4GW2FSeCOiOEvAIAA";
            if (authEnv == "dev") {
                workloadRefreshToken = "AAABAAAAo3ZCPl0FaU2WWRdLWLHpeox-FS2Q5USQ5K3oHljfFLXCqT_WL-6byZglB5MPORe88EVc1Jg6owTr-yveQl2Hi7gbFvNX4ANx0DAVX7Q5j_nMRxPqIRsE4PBvuKdRfx1vrnafyAP7WbqZ-VyoERKNM6h4yk0VgupDuaIY2AwGJ_R3jDA0aSTjhkr0iRLOL8HjvPVzUXfbUN2GlMLti0NuGKmazjg9pJZVHI61TubyFXGZGv0GiLD8Ihd8BFjaRwTtczz1-ANtQ1A2noMWKo6lTHwcQ8btoBNpqXgX68ALGkFXunBI--_okSx6i9e4V1CpyH_DjmdnrgCWy9b8GIUVRXuDNia3Lzs6d3rNZdYkxFZUvISLymh_qMFhhnyajJJEk1XRjF6D2vZIbI6Q-uqDbClNRNyVGe5CsBpDed4HnPdyyOd9b8YMJ0qztdkjDc7rJ4N5q7pjliACg_Sz8i-HWBSfFav831VSHZhJ1jNZ7bVJ3FlX-mOub2WMUdfkvXAsNjHwGXSbL3k25UJfNDW9B-EtdVMLynCZRWw0jYyvfgHkWmPCZI4oCTCfdS96ofsdIAA";
            }
        }
        else if (authEnv == "prod") {
            tokenServiceUrl = "https://login.windows.net/common/oauth2/token";
            clientId = "8563463e-ea18-4355-9297-41ff32200164";
            refreshToken = "AAABAAAAiL9Kn2Z27UubvWFPbm0gLU0qFuikw83pJJ0Sgc6bB3Ig84oewbKe7dQphka-MgTIB3w0Jxo05mYAttApqkVxRzcP1j1o0fWAp6CC0xAl1n28SuHV9CttQjR9p1i4lQYuRxT9ynGFkuTPXyxLKCMY1K5VFee_X2UvE4y1KEhTm9szkmBcreEgvoG73Fl2YhlSBnVpv-_PQqFjwPV54qCxxchYTaqduJwi7tNcD5N-pynx70HAS6DREkGy6bS_9xQNsu2FAci-CUBdSZUewuOlYKSOpHV06tr-6zikPlvAc4W69jUY6Bi-G1Ukad9B5shZx_izeuIUpGLv1T6AMR38gqRdmhNIKkchgVGEJnFoVB_jOdM7diDWN3CXjYq5MOK4BaANA8C20z4sXitAmreLpes4e8sFzcMS18KQwbmzxfvE69FzcCZtPdA_1qUq0lG2jolLlGVuEm2YOZR-FK6-arWka_Fs6OLqv0UINV1D5gT5zlBxyMMpXA-4fJDLSX5dfpXO8g_sUexCTZY2zdcAIBnH4gH9lVcaNR8Bs9AhRArXEZJCL6VOvBrkpYYK5bn6IAA";
        }
        else {
            throw "Unsupported Auth Env - '" + authEnv + "'.";
        }
        var accessToken = "";
        var agsUrlPrefix = "";
        var agsFileName = "";
        OfficeExtension.Utility._createPromiseFromResult(null).then(function () {
            agsUrlPrefix = jQuery("#" + UIConstants.TxtRichApiAgsUrl).val();
            agsUrlPrefix = OfficeExtension.Utility.trim(agsUrlPrefix);
            if (agsUrlPrefix == "") {
                throw "Missing AGS Url. It's the URL to AGS's root, such as http://shaozhu-ags:4440/zlatkom. Please replase the server name with your server name.";
            }
            if (agsUrlPrefix[agsUrlPrefix.length - 1] == '/') {
                agsUrlPrefix = agsUrlPrefix.substr(0, agsUrlPrefix.length - 1);
            }
            agsFileName = jQuery("#" + UIConstants.TxtRichApiAgsFileName).val();
            agsFileName = OfficeExtension.Utility.trim(agsFileName);
            if (agsFileName == "") {
                throw "Missing AGS file name";
            }
            jQuery("#" + UIConstants.TxtRichApiHeaderName1).val("").trigger("blur");
            jQuery("#" + UIConstants.TxtRichApiHeaderValue1).val("").trigger("blur");
            jQuery("#" + UIConstants.TxtRichApiHeaderName2).val("").trigger("blur");
            jQuery("#" + UIConstants.TxtRichApiHeaderValue2).val("").trigger("blur");
            jQuery("#" + UIConstants.TxtRichApiHeaderName3).val("").trigger("blur");
            jQuery("#" + UIConstants.TxtRichApiHeaderValue3).val("").trigger("blur");
            jQuery("#" + UIConstants.TxtRichApiRestBaseUrlOverride).val("").trigger("blur");
        }).then(function () {
            var requestInfo = {
                url: tokenServiceUrl,
                method: RichApiTest.RestUtility.httpMethodPost,
                body: "grant_type=refresh_token&refresh_token=" + refreshToken + "&client_id=" + clientId,
                headers: {
                    "CONTENT-TYPE": "application/x-www-form-urlencoded"
                }
            };
            return RichApiTest.RestUtility.invoke(requestInfo);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            var v = JSON.parse(resp.body);
            accessToken = v["access_token"];
            RichApiTest.log.comment("AccessToken=" + accessToken);
            jQuery("#" + UIConstants.TxtRichApiHeaderName1).val("Authorization").trigger("blur");
            jQuery("#" + UIConstants.TxtRichApiHeaderValue1).val("Bearer " + accessToken).trigger("blur");
        }).then(function () {
            if (workloadRefreshToken != null && workloadRefreshToken != "") {
                var requestInfo = {
                    url: tokenServiceUrl,
                    method: RichApiTest.RestUtility.httpMethodPost,
                    body: "grant_type=refresh_token&refresh_token=" + workloadRefreshToken + "&client_id=" + clientId,
                    headers: {
                        "CONTENT-TYPE": "application/x-www-form-urlencoded"
                    }
                };
                return RichApiTest.RestUtility.invoke(requestInfo);
            }
        }).then(function (resp) {
            if (workloadRefreshToken != null && workloadRefreshToken != "") {
                RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
                var v = JSON.parse(resp.body);
                accessToken = v["access_token"];
                RichApiTest.log.comment("WorkloadAccessToken=" + accessToken);
                jQuery("#" + UIConstants.TxtRichApiHeaderName3).val("Workload-Authorization").trigger("blur");
                jQuery("#" + UIConstants.TxtRichApiHeaderValue3).val("Bearer " + accessToken).trigger("blur");
            }
        }).then(function () {
            var requestInfo = {
                url: agsUrlPrefix + "/me/drive/root/children",
                method: RichApiTest.RestUtility.httpMethodGet
            };
            return RichApiTest.RestUtility.invoke(requestInfo);
        }).then(function (resp) {
            RichApiTest.RestUtility.verifyStatusCodeThrow(resp, RichApiTest.RestUtility.httpStatusOK);
            var v = JSON.parse(resp.body);
            var fileId = null;
            for (var i = 0; i < v.value.length; i++) {
                var name = v.value[i].name;
                if (OfficeExtension.Utility.caseInsensitiveCompareString(name, agsFileName)) {
                    fileId = v.value[i].id;
                    break;
                }
            }
            if (OfficeExtension.Utility.isNullOrEmptyString(fileId)) {
                throw "Cannot find file with name " + agsFileName;
            }
            RichApiTest.log.comment("FileId=" + fileId);
            var url = agsUrlPrefix + "/me/drive/items/" + fileId + "/workbook";
            RichApiTest.log.comment("URL=" + url);
            jQuery("#" + UIConstants.TxtRichApiRestBaseUrlOverride).val(url).trigger("blur");
        }).then(function () {
            RichApiTest.log.pass("Succeeded");
        }).catch(function (ex) {
            RichApiTest.log.fail(JSON.stringify(ex));
        });
    }
    RichApiTest.initGraphSettings = initGraphSettings;
})(RichApiTest || (RichApiTest = {}));
var MonacoEditorIntegration;
(function (MonacoEditorIntegration) {
    var localStorageKey = 'rich-api-test';
    var jsMonacoEditor;
    MonacoEditorIntegration.textAreaId;
    function initializeJsEditor(textAreaId, intellisensePaths) {
        MonacoEditorIntegration.textAreaId = textAreaId;
        var defaultJsText = '';
        if (window.localStorage && (localStorageKey in window.localStorage)) {
            defaultJsText = window.localStorage[localStorageKey];
        }
        var editorMode = 'text/typescript';
        jsMonacoEditor = Monaco.Editor.create(document.getElementById(textAreaId), {
            value: defaultJsText,
            mode: editorMode,
            wrappingColumn: 0,
            tabSize: 4,
            insertSpaces: false
        });
        document.getElementById(textAreaId).addEventListener('keyup', function () {
            storeCurrentJSBuffer();
        });
        if (window.location.protocol == "file:") {
            intellisensePaths = [];
        }
        else {
            intellisensePaths = intellisensePaths.map(function (path) {
                if (path.indexOf("?") < 0) {
                    path += '?';
                }
                else {
                    path += '&';
                }
                return path += 'refresh=' + Math.floor(Math.random() * 1000000000);
            });
        }
        require([
            'vs/platform/platform',
            'vs/editor/modes/modesExtensions'
        ], function (Platform, ModesExt) {
            Platform.Registry.as(ModesExt.Extensions.EditorModes).configureMode(editorMode, {
                "validate": {
                    "extraLibs": intellisensePaths
                }
            });
        });
        $(window).resize(function () {
            resizeEditor();
        });
        resizeEditor();
    }
    MonacoEditorIntegration.initializeJsEditor = initializeJsEditor;
    function getEditorValue() {
        return jsMonacoEditor.getValue();
    }
    MonacoEditorIntegration.getEditorValue = getEditorValue;
    function getEditorTextAsJavaScript() {
        var model = jsMonacoEditor.getModel();
        return model.getMode().getEmitOutput(model.getAssociatedResource(), 'js');
    }
    MonacoEditorIntegration.getEditorTextAsJavaScript = getEditorTextAsJavaScript;
    function setEditorValue(text) {
        jsMonacoEditor.getModel().setValue(text);
    }
    MonacoEditorIntegration.setEditorValue = setEditorValue;
    function resizeEditor(scrollUp) {
        if (scrollUp === void 0) { scrollUp = false; }
        $('#' + MonacoEditorIntegration.textAreaId).css('overflow', 'hidden');
        jsMonacoEditor.layout();
        if (scrollUp) {
            jsMonacoEditor.setScrollTop(0);
            jsMonacoEditor.setScrollLeft(0);
        }
        jsMonacoEditor.focus();
    }
    MonacoEditorIntegration.resizeEditor = resizeEditor;
    function storeCurrentJSBuffer() {
        if (window.localStorage) {
            window.localStorage[localStorageKey] = jsMonacoEditor.getValue();
        }
    }
})(MonacoEditorIntegration || (MonacoEditorIntegration = {}));

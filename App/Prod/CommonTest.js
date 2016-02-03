// TODO (OfficeMain: 2570648): after DTS files are published to DefinitelyTyped or are built by lab, add script to copy them into EditorIntelliSense/Published.

(function () {
	"use strict";
	
	$(document).ready(function () {
		app.initialize();

		$('#load-button').click(function () {
			var url = '//' + window.location.host + window.location.pathname + "?" +
				"endpoint=" + $('#endpoint').val() + "&" +
				"testsuite=" + $('#test-suite').val();
			window.location = url;
		})
		var endpoint = getParameterByName("endpoint").toLowerCase();
		var testsuite = getParameterByName("testsuite").toLowerCase();

		if (endpoint && testsuite) {
			$('#load-button').val("Reload");

			loadOfficeJs(endpoint, function () {
				var testParameters = getTestParameters(endpoint, testsuite);
				loadTestSuite(testParameters.endpoint, testParameters.testScripts, testParameters.testDts, testParameters.testNamespace);
			});
		}
	});

	function loadOfficeJs(endpoint, scriptLoaderOnceReady) {
		var url = getUrl();

		var options = {
			dataType: "script",
			cache: true,
			url: url
		};

		$.ajax(options).then(function () {
			Office.initialize = function () {
				$('#status').text("Loaded " + url);
				scriptLoaderOnceReady();
			};
		}).fail(function (error) {
			app.showNotification("Could not load Office.js")
		})

		function getUrl() {
			// Note: testing always against https endpoints, since they will work from both http:// and https:// protocols -- 
			//    and we had a former issue where the CDN location worked for http but not https.
			switch (endpoint) {
				case "prod":
					return 'https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js';
				case "edog":
					return 'https://eus-appsforoffice.edog.officeapps.live.com/afo/lib/1.1/hosted/office.js';
				default:
					app.showNotification("Error", "Invalid endpoint " + endpoint);
					return null;
			}
		}
	}

	function getTestParameters(endpoint, testsuite) {
		var result = {
			endpoint: "",
			testScripts: ["../RichApiTest.Core.js"],
			testDts: "",
			testNamespace: ""
		};

		result.endpoint = endpoint;

		switch (testsuite) {
			case "word":
				result.testScripts.push("../WordJsTests.js");
				result.testDts = "/Scripts/EditorIntelliSense/WordJsTests.txt";
				result.testNamespace = "WordJsTests";
				break;
			
			case "excel":
				result.testScripts.push("../ExcelTest.js");
				result.testDts = "/Scripts/EditorIntelliSense/ExcelTest.txt";
				result.testNamespace = "ExcelTest";
				break;

			default:
				app.showNotification("Error", "Invalid testsuite " + testsuite);
		}

		return result;
	}

	function loadTestSuite(endpoint, testScripts, testDts, testNamespace) {
		// For the EDOG & Prod tests, need to wait for Office to initalize to make sure that the scripts are fully loaded
		//    (e.g., otherwise, OfficeExtension.Constants.localDocumentApiPrefix is undefined)

		loadScripts(function () {
			RichApiTest.buildUI(document.getElementById("DivTests"), testNamespace, [
				"/Scripts/EditorIntelliSense/Published/OfficeCommon.d.ts",
				"/Scripts/EditorIntelliSense/Published/ExcelApi.d.ts",
				"/Scripts/EditorIntelliSense/Published/WordApi.d.ts",
				"/Scripts/EditorIntelliSense/Published/OfficeDocument.d.ts",
				"/Scripts/EditorIntelliSense/RichApiTest.Core.txt",
				"/Scripts/EditorIntelliSense/jquery.txt",
				testDts,
			]);

			$('#DivTests').show();
		});

		function loadScripts(onComplete) {
			var counter = 0;
			testScripts.forEach(function (url) {
				$.getScript(url, function () {
					counter++;
					if (counter == testScripts.length) {

						onComplete();
					}
				}).fail(function (e) {
					app.showNotification("Error retrieving file " + url, e.toString());
				})
			});
		}
	}

	function getParameterByName(name) {
		name = name.replace(/[\[]/, "\\\[").replace(/[\]]/, "\\\]");
		var regexS = "[\\?&]" + name + "=([^&#]*)";
		var regex = new RegExp(regexS);
		var results = regex.exec(window.location.search);
		if (results == null) {
			return "";
		} else {
			return decodeURIComponent(results[1]);
		}
	}

})();
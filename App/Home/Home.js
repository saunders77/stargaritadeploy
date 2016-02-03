/// <reference path="../Scripts/App.js" />
/// <reference path="../Office.Runtime.js" />
/// <reference path="../Excel.js" />
/// <reference path="../RichApiTest.Core.js" />
(function () {
	"use strict";

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
	};

	$(document).ready(function () {
		app.initialize();

		jQuery.ajaxSettings.xhr = OfficeExtension.resetXHRFactory(jQuery.ajaxSettings.xhr);
		RichApiTest.buildUI(document.getElementById("DivTests"), "ExcelTest", [
			"/Scripts/EditorIntelliSense/Excel.txt",
			"/Scripts/EditorIntelliSense/ExcelTest.txt",
			"/Scripts/EditorIntelliSense/Office.Runtime.txt",
			"/Scripts/EditorIntelliSense/RichApiTest.Core.txt",
			"/Scripts/EditorIntelliSense/Helpers.txt",
			"/Scripts/EditorIntelliSense/jquery.txt",
		]);

		$('#get-data-from-selection').click(getDataFromSelection);
		$('#getOoxmlDataFromSelection').click(getOoxmlDataFromSelection);
		$('#createPersistSession').click(function () { createSession(true); });
		$('#createNonPersistSession').click(function () { createSession(false); });
		$('#closeSessionAndClear').click(function () { closeSession(true); });
		$('#closeSession').click(function () { closeSession(false); });
	});
	

	// Reads data from current document selection and displays a notification
	function getDataFromSelection() {
		Office.context.document.getSelectedDataAsync(
			Office.CoercionType.Text,
			function (result) {
				if (result.status === Office.AsyncResultStatus.Succeeded) {
					RichApiTest.log.comment('The selected text is:' + result.value);
				} else {
					RichApiTest.log.comment('Error:' + result.error.message);
				}
			}
		);
	}

	// Reads data from current document selection and displays a notification
	function getOoxmlDataFromSelection() {
		Office.context.document.getSelectedDataAsync(
			Office.CoercionType.Ooxml,
			function (result) {
				if (result.status === Office.AsyncResultStatus.Succeeded) {
					RichApiTest.log.comment('The selected text is:' + result.value);
				} else {
					RichApiTest.log.comment('Error:' + result.error.message);
				}
			}
		);
	}

	function createSession(persistChanges) {
		ExcelTest.RestUtility.post("createsession", { persistChanges: persistChanges })
			.then(ExcelTest.RestUtility.Thenable.validateStatus(ExcelTest.RestUtility.Status.Created))
			.then(ExcelTest.RestUtility.Thenable.getBodyAsObject())
			.then(function(session) {
				RichApiTest.log.comment("persistChanges:");
				RichApiTest.log.comment(session["persistChanges"]);
				RichApiTest.log.comment("id:");
				RichApiTest.log.comment(session["id"]);
				RichApiTest.log.comment("Set TxtRichApiHeaderName2");
				jQuery("#TxtRichApiHeaderName2").val("Workbook-Session-Id").trigger("blur");
				RichApiTest.log.comment("Set TxtRichApiHeaderValue2");
				jQuery("#TxtRichApiHeaderValue2").val(session["id"]).trigger("blur");
			})
			.then(ExcelTest.pass)
			.catch(ExcelTest.reportError);
	}

	function closeSession(clearHeaderValues) {
		var request = {
			url: RichApiTest.RestUtility.getBaseUrlUsingOverride(ExcelTest.settings.baseUri) + "closeSession",
			method: RichApiTest.RestUtility.httpMethodPost
		}
		RichApiTest.RestUtility.invoke(request)
			.then(ExcelTest.RestUtility.Thenable.validateStatus(ExcelTest.RestUtility.Status.NoContent))
			.then(function () {
				if (clearHeaderValues) {
					RichApiTest.log.comment("Clear TxtRichApiHeaderName2");
					jQuery("#TxtRichApiHeaderName2").val("").trigger("blur");
					RichApiTest.log.comment("Clear TxtRichApiHeaderValue2");
					jQuery("#TxtRichApiHeaderValue2").val("").trigger("blur");
				}
			})
			.then(ExcelTest.pass)
			.catch(ExcelTest.reportError);
	}

})();

var ExcelTest;

if (!ExcelTest)
{
	ExcelTest = {};
}


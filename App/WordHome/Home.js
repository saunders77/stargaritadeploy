/// <reference path="../Scripts/App.js" />
/// <reference path="../Office.Runtime.js" />
/// <reference path="../Word.js" />
/// <reference path="../RichApiTest.Core.js" />
(function () {
	"use strict";

	// The initialize function must be run each time a new page is loaded
	Office.initialize = function (reason) {
		$(document).ready(function () {
			app.initialize();

			jQuery.ajaxSettings.xhr = OfficeExtension.resetXHRFactory(jQuery.ajaxSettings.xhr);
			RichApiTest.buildUI(document.getElementById("DivTests"), "WordJsTests", [
				"/Scripts/EditorIntelliSense/WordJsTests.txt",
				"/Scripts/EditorIntelliSense/Word.txt",
				"/Scripts/EditorIntelliSense/Office.Runtime.txt",
				"/Scripts/EditorIntelliSense/RichApiTest.Core.txt",
				"/Scripts/EditorIntelliSense/Helpers.txt",
				"/Scripts/EditorIntelliSense/jquery.txt",
			]);

			$('#get-data-from-selection').click(getDataFromSelection);
			$('#getOoxmlDataFromSelection').click(getOoxmlDataFromSelection);
		});
	};

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
})();

var WordTest;

if (!WordTest) {
	WordTest = {};
}

WordTest.reportError = function (errorInfo) {
	RichApiTest.log.comment("ErrorCode=" + errorInfo.errorCode);
	RichApiTest.log.done(false);
}

WordTest.testSetSelection= function () {
	var ctx = new Word.WordClientContext();
	var selection = ctx.document.selection;
	selection.text = "Hello World";
	ctx.executeAsync().then(function() {
		RichApiTest.log.done(true);
	}, WordTest.reportError);
}
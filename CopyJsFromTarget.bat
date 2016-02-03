:: Copy shared js files. For the runtime (but not the test framework), also copy the d.ts file for IntelliSense (renamed to .txt so that IIS will serve it)
del %~dp0App\Office.Runtime.js
mklink %~dp0App\Office.Runtime.js %TARGETROOT%\x64\debug\osfclient\x-none\Office.Runtime.js
del %~dp0App\..\Scripts\EditorIntelliSense\Office.Runtime.txt
mklink %~dp0App\..\Scripts\EditorIntelliSense\Office.Runtime.txt %TARGETROOT%\x64\debug\osfclient\x-none\Office.Runtime.d.ts
del %~dp0App\RichApiTest.Core.js
mklink %~dp0App\RichApiTest.Core.js %TARGETROOT%\x64\debug\osfclient\x-none\RichApiTest.Core.js 
del %~dp0App\..\Scripts\EditorIntelliSense\RichApiTest.Core.txt
mklink %~dp0App\..\Scripts\EditorIntelliSense\RichApiTest.Core.txt %TARGETROOT%\x64\debug\osfclient\x-none\RichApiTest.Core.d.ts

:: Copy Excel js files. Also copy the d.ts files for IntelliSense (renamed to .txt so that IIS will serve it). Do this for both the OM and tests, which can contain utility functions.
del  %~dp0App\Excel.js
mklink  %~dp0App\Excel.js %TARGETROOT%\x64\debug\xlshared\x-none\Excel.js
del %~dp0App\..\Scripts\EditorIntelliSense\Excel.txt
if exist %TARGETROOT%\x64\debug\word\x-none\WordApi.d.ts mklink %~dp0App\..\Scripts\EditorIntelliSense\Excel.txt %TARGETROOT%\x64\debug\xlshared\x-none\ExcelApi.d.ts
del %~dp0App\ExcelTest.js
if exist %TARGETROOT%\x64\debug\xlshared\x-none\ExcelTest.js mklink %~dp0App\ExcelTest.js %TARGETROOT%\x64\debug\xlshared\x-none\ExcelTest.js
del %~dp0App\..\Scripts\EditorIntelliSense\ExcelTest.txt
if exist %TARGETROOT%\x64\debug\xlshared\x-none\ExcelTest.d.ts mklink %~dp0App\..\Scripts\EditorIntelliSense\ExcelTest.txt %TARGETROOT%\x64\debug\xlshared\x-none\ExcelTest.d.ts

:: Copy the FakeExcelTest.js and its corresponding d.ts file for IntelliSense (renamed to .txt so that IIS will serve it). This is used by the FakeXlapi.html test page.
del  %~dp0App\FakeXlapiTest.js
if exist %TARGETROOT%\x64\debug\osfclient\x-none\FakeXlapiTest.js mklink  %~dp0App\FakeXlapiTest.js %TARGETROOT%\x64\debug\osfclient\x-none\FakeXlapiTest.js
del %~dp0App\..\Scripts\EditorIntelliSense\FakeXlapiTest.txt
if exist %TARGETROOT%\x64\debug\osfclient\x-none\FakeXlapiTest.d.ts mklink %~dp0App\..\Scripts\EditorIntelliSense\FakeXlapiTest.txt %TARGETROOT%\x64\debug\osfclient\x-none\FakeXlapiTest.d.ts

:: Copy Word js files. Also copy the d.ts file for IntelliSense (renamed to .txt so that IIS will serve it)
del  %~dp0App\Word.js
if exist %TARGETROOT%\x64\debug\word\x-none\Word.js mklink  %~dp0App\Word.js %TARGETROOT%\x64\debug\word\x-none\Word.js
del %~dp0App\..\Scripts\EditorIntelliSense\Word.txt
if exist %TARGETROOT%\x64\debug\word\x-none\WordApi.d.ts mklink %~dp0App\..\Scripts\EditorIntelliSense\Word.txt %TARGETROOT%\x64\debug\word\x-none\WordApi.d.ts
del %~dp0App\WordJsTests.js
if exist %TARGETROOT%\x64\debug\word\x-none\WordJsTests.js  mklink %~dp0App\WordJsTests.js %TARGETROOT%\x64\debug\word\x-none\WordJsTests.js
del %~dp0App\..\Scripts\EditorIntelliSense\WordJsTests.txt
if exist %TARGETROOT%\x64\debug\word\x-none\WordJsTests.d.ts mklink %~dp0App\..\Scripts\EditorIntelliSense\WordJsTests.txt %TARGETROOT%\x64\debug\word\x-none\WordJsTests.d.ts

:: Copy Wac files
del  %~dp0App\FakeXlapiWac.js
if exist %TARGETROOT%\x64\debug\osfclient\x-none\FakeXlapiWac.debug.js mklink  %~dp0App\FakeXlapiWac.js %TARGETROOT%\x64\debug\osfclient\x-none\FakeXlapiWac.debug.js
del %~dp0App\FakeXlapiWacTest.js
if exist %TARGETROOT%\x64\debug\osfclient\x-none\FakeXlapiWacTest.js  mklink %~dp0App\FakeXlapiWacTest.js %TARGETROOT%\x64\debug\osfclient\x-none\FakeXlapiWacTest.js 
del %~dp0App\OfficeExtension.WacRuntime.js
if exist %TARGETROOT%\x64\debug\osfclient\x-none\OfficeExtension.WacRuntime.js mklink %~dp0App\OfficeExtension.WacRuntime.js %TARGETROOT%\x64\debug\osfclient\x-none\OfficeExtension.WacRuntime.js

:: Copy Office.js related files
:: We need to copy the version specific files multiple times as the client right now expects 16.01 but WAC 16.00
SETLOCAL ENABLEDELAYEDEXPANSION
FOR /L %%A IN (0,1,1) DO (
    set officejsversion=16.0%%A
	del %~dp0Scripts\Office\1.1\excel-web-!officejsversion!.debug.js
    if exist %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Excel-Web.debug.js mklink %~dp0Scripts\Office\1.1\excel-web-!officejsversion!.debug.js %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Excel-Web.debug.js
	del %~dp0Scripts\Office\1.1\excel-web-!officejsversion!.js
    if exist %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Excel-Web.js mklink %~dp0Scripts\Office\1.1\excel-web-!officejsversion!.js %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Excel-Web.js
	del %~dp0Scripts\Office\1.1\excel-win32-!officejsversion!.debug.js
    if exist %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Excel-Win32.debug.js mklink %~dp0Scripts\Office\1.1\excel-win32-!officejsversion!.debug.js %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Excel-Win32.debug.js
	del %~dp0Scripts\Office\1.1\excel-win32-!officejsversion!.js
    if exist %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Excel-Win32.js mklink %~dp0Scripts\Office\1.1\excel-win32-!officejsversion!.js %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Excel-Win32.js
	del %~dp0Scripts\Office\1.1\Word-web-!officejsversion!.debug.js
    if exist %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Word-Web.debug.js mklink %~dp0Scripts\Office\1.1\Word-web-!officejsversion!.debug.js %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Word-Web.debug.js
	del %~dp0Scripts\Office\1.1\Word-web-!officejsversion!.js
    if exist %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Word-Web.js mklink %~dp0Scripts\Office\1.1\Word-web-!officejsversion!.js %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Word-Web.js
	del %~dp0Scripts\Office\1.1\word-win32-!officejsversion!.debug.js
    if exist %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Word-Win32.debug.js mklink %~dp0Scripts\Office\1.1\word-win32-!officejsversion!.debug.js %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Word-Win32.debug.js
	del %~dp0Scripts\Office\1.1\word-win32-!officejsversion!.js
    if exist %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Word-Win32.js mklink %~dp0Scripts\Office\1.1\word-win32-!officejsversion!.js %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Word-Win32.js
)

del %~dp0Scripts\Office\1.1\office.debug.js
if exist %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Office.debug.js mklink %~dp0Scripts\Office\1.1\office.debug.js %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Office.debug.js
del %~dp0Scripts\Office\1.1\office.js
if exist %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Office.js mklink %~dp0Scripts\Office\1.1\office.js %TARGETROOT%\x64\debug\osfweb\x-none\jscript\Office.js
del %~dp0Scripts\Office\1.1\en-us\office_strings.debug.js
if exist %TARGETROOT%\x64\debug\osfweb\en-us\jscript\office_strings.debug.js mklink %~dp0Scripts\Office\1.1\en-us\office_strings.debug.js %TARGETROOT%\x64\debug\osfweb\en-us\jscript\office_strings.debug.js
del %~dp0Scripts\Office\1.1\en-us\office_strings.js
if exist %TARGETROOT%\x64\debug\osfweb\en-us\jscript\office_strings.js mklink %~dp0Scripts\Office\1.1\en-us\office_strings.js %TARGETROOT%\x64\debug\osfweb\en-us\jscript\office_strings.js

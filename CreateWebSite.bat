iisreset
call %SystemRoot%\System32\inetsrv\appcmd delete apppool "Stargarita"
call %SystemRoot%\System32\inetsrv\appcmd delete site Stargarita
call %SystemRoot%\System32\inetsrv\appcmd add site /name:"Stargarita" /bindings:http/*:7123: /physicalPath:%~dp0
call %SystemRoot%\System32\inetsrv\appcmd unlock config "Stargarita" -section:windowsAuthentication /commit:apphost
call %SystemRoot%\System32\inetsrv\appcmd unlock config "Stargarita" -section:anonymousAuthentication /commit:apphost
call %SystemRoot%\System32\inetsrv\appcmd add apppool /name:"Stargarita"
call %SystemRoot%\System32\inetsrv\appcmd set apppool "Stargarita" /managedRuntimeVersion:v4.0
call %SystemRoot%\System32\inetsrv\appcmd set app "Stargarita/" /applicationPool:"Stargarita"
call %SystemRoot%\System32\inetsrv\appcmd set config Stargarita -section:httpProtocol /+customHeaders.[name='Access-Control-Allow-Origin',value='*'] /commit:apphost

call %~dp0devcatalog\RegisterManifests.bat

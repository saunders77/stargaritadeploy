:: Register all manifests (*.xml) files in the devcatalog folder in registry so they can be used on excel client
@setlocal enableextensions enabledelayedexpansion
@echo off
SET /A ManifestCount=0
for %%i IN (%~dp0*.xml) DO (
	SET /A ManifestCount+=1
	ECHO Registering Manifest!ManifestCount! - %~dp0%%~nxi
	REG ADD HKCU\Software\Microsoft\Office\16.0\Wef\Developer /t REG_SZ /v Manifest!ManifestCount! /d %~dp0%%~nxi /f
)
endlocal
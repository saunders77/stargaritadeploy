reg add HKEY_CURRENT_USER\Software\Microsoft\Office\16.0\Common\FeatureControl /v Enable /t REG_DWORD /d 0x301 /f
echo ^<Clearance xmlns="http://schemas.microsoft.com/office/nexus/2013/07/flightcontrol"^>^<Enabled^>^<Feature Id="35902997" /^>^</Enabled^>^<Disabled /^>^</Clearance^> > %localappdata%\Microsoft\Office\clearance.xml
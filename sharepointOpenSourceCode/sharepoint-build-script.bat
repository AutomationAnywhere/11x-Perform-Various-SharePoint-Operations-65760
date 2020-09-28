@echo on
@rem generate dll file merged with Microsoft.Sharepoint.Client dll.

set PROJECT_HOME=%~dp0
cd /d "%PROJECT_HOME%\SharepointAPI\bin\Debug"
.\ILMerge.exe /targetplatform:4.0,"C:\Windows\Microsoft.NET\Framework64\v4.0.30319" /target:library /out:"%PROJECT_HOME%\SharePointAPIWrapper.dll" Microsoft.SharePoint.Client.dll SharepointAPI.dll Microsoft.SharePoint.Client.Runtime.dll

cd %PROJECT_HOME%
pause

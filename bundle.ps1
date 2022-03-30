$msbuild="C:\Program Files (x86)\Microsoft Visual Studio\2019\Professional\MSBuild\Current\Bin\MSBuild.exe"

Remove-Item -Recurse -Force -ErrorAction SilentlyContinue ".\build"
New-Item -Path ".\build" -ItemType Directory

robocopy .\solutions\onsight_d365_connector .\build\stage01a /S
robocopy .\solutions\onsight_d365_connector_resources .\build\stage01b /S
robocopy .\solutions\onsight_d365_flow_resources .\build\stage01c /S

# Strip ".js" file extensions (D365 doesn't like them)
Get-ChildItem -Filter ".\build\stage01b\WebResources\*.js" | Rename-Item -NewName {$_.name -replace ".js","" }
Get-ChildItem -Filter ".\build\stage01b\WebResources\*.svg" | Rename-Item -NewName {$_.name -replace ".svg","" }
robocopy . .\build\stage02 [Content_Types].xml
robocopy . .\build\stage03 [Content_Types].xml input.xml *.html Onsight-32x32.png

Compress-Archive -Path .\build\stage01a\* -DestinationPath .\build\stage01a\onsight_d365_connector.zip
Compress-Archive -Path .\build\stage01b\* -DestinationPath .\build\stage01b\onsight_d365_connector_resources.zip
Compress-Archive -Path .\build\stage01c\* -DestinationPath .\build\stage01c\onsight_d365_flow_resources.zip

Invoke-Command -ScriptBlock { & "$msbuild" OnsightD365Connector.sln -p:RestorePackagesConfig=true -t:restore /property:Configuration=Release /property:Platform="Any CPU" }
Invoke-Command -ScriptBlock { & "$msbuild" OnsightD365Connector.sln /property:Configuration=Release /property:Platform="Any CPU" }
robocopy .\OnsightD365Connector\bin\Release .\build\stage02 OnsightD365Connector.dll
robocopy .\OnsightD365Connector\bin\Release\PkgFolder .\build\stage02\PkgFolder /S

Compress-Archive -Path .\build\stage02\* -DestinationPath .\build\stage03\package.zip
Compress-Archive -Path .\build\stage03\* -DestinationPath .\build\Onsight_Dynamics_365_Field_Service_Connector.zip

Remove-Item -Recurse -Force -ErrorAction SilentlyContinue ".\build"
New-Item -Path ".\build" -ItemType Directory

robocopy .\solutions\onsight_d365_connector .\build\stage01a /S
robocopy .\solutions\onsight_d365_connector_resources .\build\stage01b /S
# Strip ".js" file extensions (D365 doesn't like them)
Get-ChildItem -Filter ".\build\stage01b\WebResources\*.js" | Rename-Item -NewName {$_.name -replace ".js","" }
Get-ChildItem -Filter ".\build\stage01b\WebResources\*.svg" | Rename-Item -NewName {$_.name -replace ".svg","" }
robocopy . .\build\stage02 [Context_Types].xml
New-Item -Path ".\build\stage02\PkgFolder" -ItemType Directory
robocopy . .\build\stage03 [Context_Types].xml input.xml *.html Onsight-32x32.png

Compress-Archive -Path .\build\stage01a\* -DestinationPath .\build\stage02\PkgFolder\onsight_d365_connector.zip
Compress-Archive -Path .\build\stage01b\* -DestinationPath .\build\stage02\PkgFolder\onsight_d365_connector_resources.zip
Compress-Archive -Path .\build\stage02\* -DestinationPath .\build\stage03\package.zip
Compress-Archive -Path .\build\stage03\* -DestinationPath .\build\Onsight_Dynamics_365_Field_Service_Connector.zip

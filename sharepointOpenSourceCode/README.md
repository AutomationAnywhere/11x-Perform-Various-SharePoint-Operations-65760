# Sharepoint Metabot

DLL project to call sharepoint API using CSOM (Client Side Object Model)

## How to make a dll

1. Clone this repository
2. Open this solution named ConsoleApp2.sln using the Visual Studio (Express or Enterprise).
3. Build this solution.
4. Run ``sharepoint-build-script.bat``
5. Then, you can get the ``SharepointAPIWrapper.dll`` under this folder.

## How to test a code

1. You need to prepare your SharePoint access.
2. Open the solution using the Visual Studio (Express or Enterprise).
3. Open the Resource4Test.resx under SharePointAPIWrapper.Test project.
4. Set your credential, valid SharePoint URL for files and folders.
5. Then, you can run unit tests in the Visual Studio (The easiest way is run test code using right click on the method name).



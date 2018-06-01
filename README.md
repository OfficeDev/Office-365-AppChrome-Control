# Office 365 AppChrome Control

========================

The Office 365 AppChrome Control provides a simple but extensible way to add Office 365-style navigation bar with your own functions. This enables users to sign in and out of their Office 365 account and navigate to sites and pages that you'd like to feature. We develop the control in JavaScript to provide universal compatibility without the additional overhead of other frameworks. We build the controls with two parts, Web UI and Data Provider, so that developers could customize the control easily based on the interface we defined.

You could find more description from - [Office 365 JavaScript controls](https://msdn.microsoft.com/en-us/office/office365/howto/javascript-controls) 

## Web UI
The standard Office 365 web experience comes from Office UI Fabric, you could also visit their GitHub repository - [OfficeDev/Office-UI-Fabric](https://github.com/OfficeDev/Office-UI-Fabric)

## Data Provider
We provide a sample data provider which retrieves data from Office 365, you could get more detail about how to access Office 365 data from - [Office 365 API reference](https://msdn.microsoft.com/office/office365/HowTo/rest-api-overview). If you want to use the sample provider, please remember to type in your Office 365 client ID before initialize the provider. 

Here are the key steps
 - Create your own Office 365 client ID
 - Set permissions for the API you plan to use
 - Set "Redirect URI" for your web app
 - Configure "Implicit Grant" for OAuth flow

For more detail guidance, you could check from - [Create an app with Office 365 APIs](https://msdn.microsoft.com/office/office365/howto/getting-started-Office-365-APIs)

## Permissions 
You need to configure permissions for your Office 365 app based on the API and scope you want to access. 

Here are the permissions sample data provider requires

|Feature|Application Name|Delegated Permission|
|:-----|:-----|:-----|
|Login|Azure Active Directory|Sign in and read user profile|
|User's info|Azure Active Directory|Sign in and read user profile|

## License
 - All files on the Office 365 AppChrome Control repository are subject to the MIT license. Please read the License file at the root of the project. 
 - All the Web UI are based on [OfficeDev/Office-UI-Fabric](https://github.com/OfficeDev/Office-UI-Fabric)
 - Usage of the fonts referenced on Office UI Fabric files is subject to the terms listed here 

## How to store access token 
We provide a sample implicit grant login provider which is based on ADAL.js. The ADAL.js stores the access token in browser's localStorage. The use of localStorage has security implications, given that other apps in the same domain will have access to it, and it is prone to the same attacks that localStorage have to deal with. So before using access token to do some user credential related operation, it must be sent to a backend service for parsing and validating. 

We provide server side sample project LoginControlForSPASolutionabout to validate access token. More ways about how to validate access token, please reference here: https://github.com/AzureADSamples.

## Sample Site
We provide a sample site in the "demo" folder. In this site, you could find
 - Demo for the control
 - Test page

Here are the key steps for running your own sample site

install nodejs
 - http://nodejs.org

To install development packages - From the root of your local git repository
 - npm install gulp -g
 - npm install
 - npm install dev

To build minified files
 - gulp
Then you could deploy to your web service


This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/). For more information, see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.

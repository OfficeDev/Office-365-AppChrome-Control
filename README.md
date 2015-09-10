# Office 365 AppChrome Control

========================

The Office 365 AppChrome - With the AppChrome control, your app can show an Office 365-style navigation bar as well as customized links. This enables users to sign in and out of their Office 365 account and navigate to sites and pages that you'd like to feature.
This control is a pure front-end control. It only contains JS, HTML, CSS. 

## Setup dev machine
Setup Git
* https://help.github.com/articles/set-up-git/

Clone this repository to a local repository
* https://help.github.com/articles/cloning-a-repository/

install nodejs
* http://nodejs.org

To install development packages - From the root of your local git repository
* npm install

To build minified files
* gulp

## Note
Implicit grant login provider implementation is based on ADAL.js, which stores the access token in browser's localStorage. The use of localStorage has security implications, given that other apps in the same domain will have access to it, and it is prone to all the same attacks that localStorage have to deal with. So before using access token to do some user credential related operation, it must be sent to a backend service for parsing and validating. We provide server side sample project LoginControlForSPASolutionabout to validat access token. More ways about how to validate access token, please reference here: https://github.com/AzureADSamples




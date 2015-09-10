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
if using the implicit grant login provider we provided, it has security issue if user uses access token directly to do some user credential related operation. The access token must be sent to a backend service for parsing and validating. And we provide server side validating token sample code. Please reference the sample project LoginControlForSPASolution.






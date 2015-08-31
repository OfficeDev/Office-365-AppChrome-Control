var element = document.getElementById("container");

var login = new Office.Controls.ImplicitGrantLogin({
    clientId: '********-****-****-****-************',
    redirectUri: window.location.href,
    postLogoutRedirectUri: window.location.toString(),
    cacheLocation: 'localStorage', // enable this for IE, as sessionStorage does not work for localhost.
});

var appChrome = new Office.Controls.AppChrome("AppName", element, login, {});

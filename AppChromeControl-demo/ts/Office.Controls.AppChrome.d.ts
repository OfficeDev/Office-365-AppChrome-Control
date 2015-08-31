declare module Office.Controls {
    interface LoginConfig {
        clientId: string;
        redirectUri: string;
        postLogoutRedirectUri: string;
        cacheLocation: string;
    }
    class ImplicitGrantLogin {
        constructor(config: Office.Controls.LoginConfig);
    }
    class AppChrome {
        constructor(appName: string, root: HTMLElement, loginProvider: ImplicitGrantLogin, options);
    }
}
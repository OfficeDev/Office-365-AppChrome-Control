(function () {
    "use strict";

    if (window.Type && window.Type.registerNamespace) {
        Type.registerNamespace('Office.Controls');
    } else {
        if (window.Office === undefined) {
            window.Office = {}; window.Office.namespace = true;
        }
        if (window.Office.Controls === undefined) {
            window.Office.Controls = {}; window.Office.Controls.namespace = true;
        }
    }

    Office.Controls.AppChrome = function (root, loginProvider, options) {
        if (typeof root !== 'object' || typeof loginProvider !== 'object' || (!Office.Controls.Utils.isNullOrUndefined(options) && typeof options !== 'object')) {
            Office.Controls.Utils.errorConsole('Invalid parameters type');
            return;
        }
        this.rootNode = root;
        this.loginProvider = loginProvider;
        if (!Office.Controls.Utils.isNullOrUndefined(options)) {
            if (!Office.Controls.Utils.isNullOrUndefined(options.appTitle)) {
                this.appDisPlayName = options.appTitle;
            }
            if (!Office.Controls.Utils.isNullOrUndefined(options.appURI)) {
                this.appURI = options.appURI;
            }
            if (!Office.Controls.Utils.isNullOrUndefined(options.settingsLinks)) {
                this.settingsLinks = options.settingsLinks;
            }
        }
        this.renderControl();
    };

    Office.Controls.AppChrome.prototype = {
        rootNode: null,
        dropDownListNode: null,
        loginProvider: null,
        appDisPlayName: null,
        appURI: null,
        settingsLinks: null,
      

        renderControl: function () {
            this.rootNode.innerHTML = Office.Controls.appChromeTemplates.generateBannerTemplate(this.appDisPlayName, this.appURI, this.loginProvider);
            var dropDonwListRoot = document.createElement("div");
            dropDonwListRoot.innerHTML = Office.Controls.appChromeTemplates.generateDropDownList(this.settingsLinks);
            this.rootNode.parentNode.insertBefore(dropDonwListRoot, this.rootNode.nextSibling);
        },


    };

    Office.Controls.appChromeTemplates = function () {
    };

    Office.Controls.appChromeTemplates.generateBannerTemplate = function (appDisPlayName, appURI, loginProvider) {
        var body = '<div id=\"GeminiShellHeader\" class=\"removeFocusOutline\"><div autoid=\"_o365sg2c_k\" class=\"o365cs-nav-header16 o365cs-base o365cst o365spo o365cs-nav-header o365cs-topnavBGColor-2 o365cs-topnavBGImage\" id="O365_NavHeader\">';
        body += Office.Controls.appChromeTemplates.generateLeftPart(appDisPlayName, appURI);
        body += Office.Controls.appChromeTemplates.generateMiddlePart();
        body += Office.Controls.appChromeTemplates.generateRightPart(loginProvider);
        body += '</div></div>';
        return body;
    };

    Office.Controls.appChromeTemplates.generateLeftPart = function (appDisPlayName, appURI) {
        var innerHtml = '<div class=\"o365cs-nav-leftAlign\"><div class=\"o365cs-nav-topItem\"><button type=\"button\" class=\"o365cs-nav-item o365cs-nav-button o365cs-navMenuButton ms-bgc-tdr-h o365button ms-bgc-tp\" role=\"menuitem\" id=\"O365_MainLink_NavMenu\" aria-label=\"Open the app launcher to access your Office 365 apps\">';
        innerHtml += '<div class=\"o365cs-base o365cst o365cs-nav-navMenu popupShadow removeFocusOutline\" ispopup=\"1\" tabindex=\"0\" style=\"display: none;\"></div>';
        innerHtml +='<div class=\"o365cs-base o365cst o365cs-nav-inactivityCallout popupShadow removeFocusOutline\" ispopup=\"1\" tabindex=\"0\" style=\"display: none;\"></div>';
        innerHtml +='<span class=\"wf wf-size-x18 wf-family-o365\" role=\"presentation\">Home</span><div class=\"o365cs-nav-navMenuBeak\" style=\"display: none;\"></div>';
        innerHtml +='<div class=\"o365cs-nav-inactivityCalloutBeak ms-bcl-tp\" style=\"display: none;\"></div></button></div>';
        innerHtml +='<div class=\"o365cs-nav-topItem ms-fcl-w\" style=\"display: none;\"></div>';
        innerHtml +='<div class=\"o365cs-nav-topItem o365cs-nav-o365Branding\">';
        innerHtml +='<a class=\"o365cs-nav-bposLogo o365cs-topnavText o365cs-o365logo o365button\" role=\"link\" id=\"O365_MainLink_Logo\" href=\"http://portal.office.com\" aria-label=\"Go to your Office 365 home page\"><span class=\"o365cs-nav-brandingText\">Office 365</span></a>';
        innerHtml +='<div class=\"o365cs-nav-appTitleLine o365cs-nav-brandingText o365cs-topnavText\"></div>';
        innerHtml += '<a class=\"o365cs-nav-appTitle o365cs-topnavText o365button\" role=\"link\" href=\"https://portal.office.com/admin/default.aspx\" aria-label=\"Go to the Office 365 admin center\">';
        innerHtml += '<span class=\"o365cs-nav-brandingText\" id=\"change_name\">' + Office.Controls.Utils.htmlEncode(appDisPlayName) + '</a></div><div class=\"o365cs-nav-topItem o365cs-breadCrumbContainer\" style=\"display: none;\"></div></div>'
        return innerHtml;
    };

    Office.Controls.appChromeTemplates.generateMiddlePart = function () {
        var innerHtml = '<div class=\"o365cs-nav-centerAlign\"><div style=\"display: none;\"></div><div style=\"display: none;\"></div></div>';
        return innerHtml;
    };

    Office.Controls.appChromeTemplates.generateRightPart = function (loginProvider) {
        if (Office.Controls.Utils.isNullOrUndefined(loginProvider.hasSignedIn) || loginProvider.hasSignedIn()== false) {
            var innerHtml = '<div class=\"o365cs-nav-rightAlign o365cs-topnavLinkBackground-2\" id=\"O365_TopMenu\"><div><div class=\"o365cs-nav-headerRegion\"><div class=\"o365cs-nav-notificationTrayContainer\"><div class=\"o365cs-w100-h100" style="display: none;\"></div></div>';
            innerHtml += '<div class=\"o365cs-nav-pinnedAppsContainer\"><div class=\"o365cs-nav-pinnedApps\"><div></div></div></div></div><div class=\"o365cs-nav-rightMenus\"><div role=\"banner\" aria-label=\"User settings\">';
            innerHtml += '<div class=\"o365cs-nav-topItem\"><button autoid=\"_o365sg2c_0\" type=\"button\" class=\"o365cs-nav-item o365cs-nav-button ms-fcl-w o365cs-me-nav-item o365button ms-bgc-tdr-h\" role=\"menuitem\" aria-label=\"offline menu with submenu\" aria-haspopup=\"true\" id=\"login_user\">';
            innerHtml += '<div class=\"o365cs-me-tileview-container\"><div autoid=\"_o365sg2c_1\" class=\"o365cs-me-tileview\"><div autoid=\"_o365sg2c_2\" class=\"o365cs-me-presence5x50 o365cs-me-presenceColor-Offline\"></div>';
            innerHtml += '<span autoid=\"_o365sg2c_3\" class=\"ms-bgc-nt ms-fcl-w o365cs-me-tileimg o365cs-me-tileimg-doughboy owaimg wf wf-size-x52 wf-o365-people wf-family-o365\" role=\"presentation\"></span>';
            innerHtml += '<div style=\"display: none;\"></div><div class=\"o365cs-me-tileimg\"><img autoid=\"_o365sg2c_5\" class=\"o365cs-me-personaimg\" src=\"image/default.jpg\" style=\"display: inline; width: 50px; top: 0px;\" id=\"login_user_image\"></div></div></div>';
            innerHtml += '<div><div autoid=\"_o365sg2c_6\" class=\"o365cs-me-tile-nophoto\"><div autoid=\"_o365sg2c_7\" class=\"o365cs-me-presenceOffline5x50\"></div>';
            innerHtml += '<div class=\"o365cs-me-tile-nophoto-username-container\"><span autoid=\"_o365sg2c_8\" class=\"o365cs-me-tile-nophoto-username o365cs-me-bidi\" id=\"UserName\" style=\"display:none\"></span></div>';
            innerHtml += '<span class=\"wf-o365-x18 ms-fcl-nt o365cs-me-tile-nophoto-down owaimg wf wf-size-x18 wf-o365-downcarat wf-family-o365\" role=\"presentation\"></span></div></div></button></div></div>';
            innerHtml += '<div class=\"o365cs-w100-h100\"><div><div class=\"o365cs-notifications-notificationPopupArea o365cs o365cs-base o365cst\" ispopup=\"1\" style=\"display: none;\"></div><div style=\"display: none;\"></div></div></div></div></div></div>';
            return innerHtml;
        }
        
    };

    Office.Controls.appChromeTemplates.generateDropDownList = function (appLinks) {
        var innerHtml = '<div class=\"o365cs-nav-contextMenu o365spo contextMenuPopup removeFocusOutline\" ispopup=\"1\" iscontextmenu=\"1\" role=\"menu\" ismodal=\"false\" tabindex=\"-1\" parentids=\"(6)\" style=\"min-width: 150px; position: absolute; box-sizing: border-box; outline: 0px; z-index: 2003; right: 0px; top: 60px; display: none;\" id=\"_ariaId_7\">';
        innerHtml += '<div class=\"o365cs-base ms-bgc-w o365cst o365cs-context-font o365cs-me-contextMenu\"><div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_0\" class=\"o365cs-me-itemsList\" tabindex=\"-1\" id=\"additem\"><div>';
        innerHtml += Office.Controls.appChromeTemplates.generatePersonPart();
        if (!Office.Controls.Utils.isNullOrUndefined(appLinks)) {
            for (var name in appLinks) {
                innerHtml += Office.Controls.appChromeTemplates.generateAppLinkPart(name, appLinks[name]);
            }
        }
        innerHtml += Office.Controls.appChromeTemplates.generateSignOutPart();
        innerHtml += '</div></div></div></div>';
        return innerHtml;
    };

    Office.Controls.appChromeTemplates.generatePersonPart = function () {
        var innerHtml = '<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" tabindex=\"-1\" aria-selected=\"false\"><div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_2\" class=\"o365cs-userInfo o365button\" role=\"group\" tabindex=\"0\"><div class=\"o365cs-me-persona\"><div class=\"o365cs-me-personaView\" id=\"myPersona\"></div></div></div>';
        innerHtml += '<div class=\"o365button\" role=\"menuitem\" tabindex=\"0\"><div style=\"display: none;\"></div></div>';
        innerHtml += '<div class=\"o365button o365cs-contextMenuItem ms-fcl-b ms-bgc-nl-h\" role=\"menuitem\" tabindex=\"0\" aria-label=\"Sign in to add another account\" title=\"Sign in to add another account\" style=\"display: none;\"></div><div><div><div class=\"_fce_p ms-bcl-nl\"></div></div></div></div>';
        return innerHtml;
    };

    Office.Controls.appChromeTemplates.generateAppLinkPart = function (name, link) {
        var innerHtml = '<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" tabindex=\"-1\"><div class=\"o365cs-contextMenuSeparator ms-bcl-nl\"></div></div>';
        innerHtml += '<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" tabindex=\"-1\" aria-selected=\"false\"><a class=\"o365button o365cs-contextMenuItem ms-fcl-b\" role=\"link\" href=\"'+link+'\">';
        innerHtml += '<div class=\"_fce_j\"><span class=\"_fce_k owaimg\" role=\"presentation\" style=\"display: none;\"></span><span autoid=\"_fce_4\">' + Office.Controls.Utils.htmlEncode(name) + '</span></div></a></div>';
        return innerHtml;
    };

    Office.Controls.appChromeTemplates.generateSignOutPart = function () {
        var innerHtml = '<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" tabindex=\"-1\"><div class=\"o365cs-contextMenuSeparator ms-bcl-nl\"></div></div>';
        innerHtml += '<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" tabindex=\"-1\" aria-label=\"Sign out and return to the Sign-in page\" title=\"Sign out and return to the Sign-in page\" aria-selected=\"false\"><a class=\"o365button o365cs-contextMenuItem ms-fcl-b\" role=\"link\" id=\"O365_SubLink_ShellSignout\" aria-label=\"Sign out of Office 365 and return to the Sign-in page\">';
        innerHtml += '<div class=\"_fce_j\"><span class=\"_fce_k owaimg\" role=\"presentation\" style=\"display: none;\"></span><span autoid=\"_fce_4\" aria-label=\"Sign out of Office 365 and return to the Sign-in page\">Sign out</span></div></a></div>';
        return innerHtml;
    };


})();

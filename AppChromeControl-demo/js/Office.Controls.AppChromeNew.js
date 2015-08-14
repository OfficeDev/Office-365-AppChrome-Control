(function() {
    "use strict";

    if (window.Type && window.Type.registerNamespace) {
        Type.registerNamespace('Office.Controls');
    } else {
        if (window.Office === undefined) {
            window.Office = {};
            window.Office.namespace = true;
        }
        if (window.Office.Controls === undefined) {
            window.Office.Controls = {};
            window.Office.Controls.namespace = true;
        }
    }

    Office.Controls.AppChrome = function(root, loginProvider, options) {
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
        if (!Office.Controls.Utils.isNullOrUndefined(loginProvider.hasSignedIn)) {
            this.isSignedIn = loginProvider.hasSignedIn();
        }
        this.registerinnerText();
        this.renderControl();
        if (this.isSignedIn == true) {
            var instance = this;
            loginProvider.getUserInfoAsync(function(error, userData) {
                if (!Office.Controls.Utils.isNullOrUndefined(userData)) {
                    instance.signedUserInfo = userData;
                } else {
                    instance.isSignedIn = false;
                    Office.Controls.Utils.errorConsole('Getting User info failed');
                }
                instance.updateControl();
            });
        } else {
            this.updateControl();
        }
    };

    Office.Controls.AppChrome.prototype = {
        rootNode: null,
        dropDownListNode: null,
        loginProvider: null,
        appDisPlayName: null,
        appURI: null,
        settingsLinks: null,
        isSignedIn: false,
        signedUserInfo: null,
        defaultImage: "data:image/jpeg;base64,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",

        registerinnerText: function() {
            var lBrowser = {};
            lBrowser.isW3C = document.getElementById ? true : false;
            lBrowser.isNS6 = lBrowser.isW3C && (navigator.appName == "Netscape");
            if (lBrowser.isNS6) { //firefox innerText define   
                HTMLElement.prototype.__defineGetter__("innerText", function() {
                    return this.textContent;
                });
                HTMLElement.prototype.__defineSetter__("innerText", function(sText) {
                    this.textContent = sText;
                });
            }
        },

        renderControl: function() {
            this.rootNode.innerHTML = Office.Controls.appChromeTemplates.generateBannerTemplate(this.appDisPlayName, this.appURI);
            var dropDonwListRoot = document.createElement("div");
            dropDonwListRoot.innerHTML = Office.Controls.appChromeTemplates.generateDropDownList(this.settingsLinks);
            this.rootNode.parentNode.insertBefore(dropDonwListRoot, this.rootNode.nextSibling);
            var instance = this;
            document.getElementById('O365_SubLink_ShellSignout').addEventListener('click', function() {
                instance.loginProvider.signOut();
            });
        },

        updateControl: function() {
            var instance = this;
            var loginButton = document.getElementById('login_user');
            var Personalistview = document.getElementById('_ariaId_7');
            if (this.isSignedIn == false) {
                document.getElementById('dropdownIcon').style.display = 'none';
                document.getElementById('image_container').style.display = 'none';
                document.getElementById('user_name').innerText = Office.Controls.Utils.htmlEncode("Sign In");
                loginButton.addEventListener('click', function() {
                    instance.loginProvider.signIn();
                });
            } else {
                document.getElementById('user_name').innerText = Office.Controls.Utils.htmlEncode(this.signedUserInfo.displayName);
                if (this.signedUserInfo.imgSrc != null) {
                    document.getElementById('login_user_image').src = this.signedUserInfo.imgSrc;
                } else {
                    document.getElementById('login_user_image').src = this.defaultImage;
                }
                document.getElementById('login_user_image').title = this.signedUserInfo.displayName;
                this.genInlinePersona(document.getElementById('myPersona'));
                loginButton.addEventListener('click', function() {
                    if (Personalistview.style.display == 'none') {
                        Personalistview.style.display = 'block';
                    } else {
                        Personalistview.style.display = 'none';
                    }
                    instance.changeTopMenuColor();
                });
                document.onclick = function(e) {
                    if (Personalistview.style.display == 'block') {
                        e = e || event;
                        var target = e.target || e.srcElement;
                        while (target) {
                            if (target == loginButton || target == Personalistview) {
                                Personalistview.style.display = 'block';
                                break;
                            } else {
                                Personalistview.style.display = 'none';
                            }
                            target = target.parentNode;
                        }
                    }
                    instance.changeTopMenuColor();
                }
            }
        },

        changeTopMenuColor: function() {
            var Personalistview = document.getElementById('_ariaId_7');
            if (Personalistview.style.display == 'block') {
                document.getElementById('O365_TopMenu').style.backgroundColor = '#fff';
                document.getElementById('user_name').style.color = '#000';
            } else {
                document.getElementById('O365_TopMenu').style.backgroundColor = '#005a9e';
                document.getElementById('user_name').style.color = '#fff';
            }
        },

        genInlinePersona: function(ele) {
            if (typeof ele !== 'object') {
                Office.Controls.Utils.errorConsole('Invalid parameters type');
                return;
            }
            var personaObj = {};
            personaObj.SecondaryText = this.signedUserInfo.accountName;
            personaObj.PrimaryText = this.signedUserInfo.displayName;
            personaObj.ImageUrl = this.defaultImage;
            if (this.signedUserInfo.imgSrc != null) {
                personaObj.ImageUrl = this.signedUserInfo.imgSrc;
            }
            Office.Controls.Persona.PersonaHelper.createInlinePersona(ele, personaObj);
        }
    };

    Office.Controls.appChromeTemplates = function() {};

    Office.Controls.appChromeTemplates.generateBannerTemplate = function(appDisPlayName, appURI) {
        var body = '<div id=\"GeminiShellHeader\" class=\"removeFocusOutline\"><div autoid=\"_o365sg2c_k\" class=\"o365cs-nav-header16 o365cs-base o365cs-topnavBGColor-2 o365cs-topnavBGImage\" id="O365_NavHeader\">';
        body += Office.Controls.appChromeTemplates.generateLeftPart(appDisPlayName, appURI);
        body += Office.Controls.appChromeTemplates.generateMiddlePart();
        body += Office.Controls.appChromeTemplates.generateRightPart();
        body += '</div></div>';
        return body;
    };

    Office.Controls.appChromeTemplates.generateLeftPart = function(appDisPlayName, appURI) {
        var innerHtml = '<div class=\"o365cs-nav-leftAlign\">';
        innerHtml += '<div class=\"o365cs-nav-topItem o365cs-nav-o365Branding\">';
        innerHtml += '<a class=\"o365cs-nav-bposLogo o365cs-topnavText o365cs-o365logo o365button\" role=\"link\" id=\"O365_MainLink_Logo\" href=\"http://portal.office.com\" aria-label=\"Go to your Office 365 home page\"><span class=\"o365cs-nav-brandingText\">Office 365</span></a>';
        innerHtml += '<div class=\"o365cs-nav-appTitleLine o365cs-nav-brandingText o365cs-topnavText\"></div>';
        innerHtml += '<a class=\"o365cs-nav-appTitle o365cs-topnavText o365button\" role=\"link\" href=\"' + appURI + '\" aria-label=\"Go to the App home page\">';
        innerHtml += '<span class=\"o365cs-nav-brandingText\" id=\"change_name\">' + Office.Controls.Utils.htmlEncode(appDisPlayName) + '</a></div></div>'
        return innerHtml;
    };

    Office.Controls.appChromeTemplates.generateMiddlePart = function() {
        var innerHtml = '<div class=\"o365cs-nav-centerAlign\"></div>';
        return innerHtml;
    };

    Office.Controls.appChromeTemplates.generateRightPart = function() {
        var innerHtml = '<div class=\"o365cs-nav-rightAlign o365cs-topnavLinkBackground-2\" id=\"O365_TopMenu\"><div>';
        innerHtml += '<div class=\"o365cs-nav-rightMenus\"><div role=\"banner\" aria-label=\"User settings\">';
        innerHtml += '<div class=\"o365cs-nav-topItem\"><button autoid=\"_o365sg2c_0\" type=\"button\" class=\"o365cs-nav-item o365cs-nav-button ms-fcl-w o365cs-me-nav-item o365button ms-bgc-tdr-h\" role=\"menuitem\" aria-label=\"offline menu with submenu\" aria-haspopup=\"true\" id=\"login_user\">';

        innerHtml += '<div class=\"o365cs-me-tileview-container\" id=\"image_container\">';

        innerHtml += '<div autoid=\"_o365sg2c_1\" class=\"o365cs-me-tileview\"><div class=\"o365cs-me-tileimg\"><img autoid=\"_o365sg2c_5\" class=\"o365cs-me-personaimg\" src=\"image/default.jpg\" style=\"display: inline; width: 50px; top: 0px;\" id=\"login_user_image\"></div></div></div>';

        innerHtml += '<div class=\"o365cs-me-tile-container\"><div autoid=\"_o365sg2c_6\" class=\"o365cs-me-tile-nophoto\"><div class=\"o365cs-me-tile-nophoto-username-container\">';

        innerHtml += '<span autoid=\"_o365sg2c_8\" class=\"o365cs-me-tile-nophoto-username o365cs-me-bidi\" id=\"user_name\"></span></div>';
        innerHtml += '<span class=\"wf-o365-x18 ms-fcl-nt o365cs-me-tile-nophoto-down owaimg wf wf-size-x18 ms-Icon--caretDown wf-family-o365\" role=\"presentation\" style=\"display:table-cell\" id=\"dropdownIcon\"></span></div></div>'
        innerHtml += '</button></div></div></div></div></div>';
        return innerHtml;


    };

    Office.Controls.appChromeTemplates.generateDropDownList = function(appLinks) {
        var innerHtml = '<div class=\"o365cs-nav-contextMenu o365spo contextMenuPopup removeFocusOutline\" ispopup=\"1\" iscontextmenu=\"1\" role=\"menu\" ismodal=\"false\" tabindex=\"-1\" parentids=\"(6)\" style=\"min-width: 150px; position: absolute; box-sizing: border-box; outline: 0px; z-index: 2003; right: 0px; top: 60px; display: none;\" id=\"_ariaId_7\">';
        innerHtml += '<div class=\"o365cs-base ms-bgc-w o365cst o365cs-context-font o365cs-me-contextMenu\"><div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_0\" class=\"o365cs-me-itemsList\" tabindex=\"-1\" id=\"additem\"><div>';
        innerHtml += Office.Controls.appChromeTemplates.generatePersonaPart();
        if (!Office.Controls.Utils.isNullOrUndefined(appLinks)) {
            for (var name in appLinks) {
                innerHtml += Office.Controls.appChromeTemplates.generateAppLinkPart(name, appLinks[name]);
            }
        }
        innerHtml += Office.Controls.appChromeTemplates.generateSignOutPart();
        innerHtml += '</div></div></div></div>';
        return innerHtml;
    };

    Office.Controls.appChromeTemplates.generatePersonaPart = function() {
        var innerHtml = '<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" tabindex=\"-1\" aria-selected=\"false\"><div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_2\" class=\"o365cs-userInfo o365button\" role=\"group\" tabindex=\"0\"><div class=\"o365cs-me-persona\"><div class=\"o365cs-me-personaView\" id=\"myPersona\"></div></div></div>';
        innerHtml += '<div class=\"o365button\" role=\"menuitem\" tabindex=\"0\"><div style=\"display: none;\"></div></div>';
        innerHtml += '<div class=\"o365button o365cs-contextMenuItem ms-fcl-b ms-bgc-nl-h\" role=\"menuitem\" tabindex=\"0\" aria-label=\"Sign in to add another account\" title=\"Sign in to add another account\" style=\"display: none;\"></div><div><div><div class=\"_fce_p ms-bcl-nl\"></div></div></div></div>';
        innerHtml += '<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" tabindex=\"-1\"><div class=\"o365cs-contextMenuSeparator ms-bcl-nl\"></div></div>'
        return innerHtml;
    };

    Office.Controls.appChromeTemplates.generateAppLinkPart = function(name, link) {
        var innerHtml = '<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" tabindex=\"-1\" aria-selected=\"false\"><a class=\"o365button o365cs-contextMenuItem ms-fcl-b\" role=\"link\" href=\"' + link + '\">';
        innerHtml += '<div class=\"_fce_j\"><span class=\"_fce_k owaimg\" role=\"presentation\" style=\"display: none;\"></span><span autoid=\"_fce_4\">' + Office.Controls.Utils.htmlEncode(name) + '</span></div></a></div>';
        return innerHtml;
    };

    Office.Controls.appChromeTemplates.generateSignOutPart = function() {
        var innerHtml = '<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" tabindex=\"-1\"><div class=\"o365cs-contextMenuSeparator ms-bcl-nl\"></div></div>';
        innerHtml += '<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" tabindex=\"-1\" aria-label=\"Sign out and return to the Sign-in page\" title=\"Sign out and return to the Sign-in page\" aria-selected=\"false\"><a class=\"o365button o365cs-contextMenuItem ms-fcl-b\" role=\"link\" id=\"O365_SubLink_ShellSignout\" aria-label=\"Sign out of Office 365 and return to the Sign-in page\">';
        innerHtml += '<div class=\"_fce_j\"><span class=\"_fce_k owaimg\" role=\"presentation\" style=\"display: none;\"></span><span autoid=\"_fce_4\" aria-label=\"Sign out of Office 365 and return to the Sign-in page\">Sign out</span></div></a></div>';
        return innerHtml;
    };


})();
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

    function isEmpty(testString) {
        if (testString.replace(/(^\s+)|(\s+$)/g, "").length != 0) {
            return false;
        } else {
            return true;
        }
    }

    function ValidUrl(testString) {
        var pattern = new RegExp('^(https?:\\/\\/)?' + // protocol
            '((([a-z\\d]([a-z\\d-]*[a-z\\d])*)\\.)+[a-z]{2,}|' + // domain name
            '((\\d{1,3}\\.){3}\\d{1,3}))' + // OR ip (v4) address
            '(\\:\\d+)?(\\/[-a-z\\d%_.~+]*)*' + // port and path
            '(\\?[;&a-z\\d%_.~+=-]*)?' + // query string
            '(\\#[-a-z\\d_]*)?$', 'i'); // fragment locator
        if (!pattern.test(testString)) {
            return false;
        } else {
            return true;
        }
    }

    Office.Controls.AppChrome = function(appTitle, root, loginProvider, options) {
        if (typeof root !== 'object' || typeof loginProvider !== 'object' || (!Office.Controls.Utils.isNullOrUndefined(options) && typeof options !== 'object')) {
            Office.Controls.Utils.errorConsole('Invalid parameters type');
            return;
        }
        this.rootNode = root;
        this.loginProvider = loginProvider;
        if (!Office.Controls.Utils.isNullOrUndefined(appTitle) && !isEmpty(appTitle)) {
            if (appTitle.length >= 50) {
                appTitle = appTitle.substr(0, 50);
            }
            this.appDisPlayName = appTitle;
        }
        if (!Office.Controls.Utils.isNullOrUndefined(options)) {
            if (!Office.Controls.Utils.isNullOrUndefined(options.appHomeUrl) && !isEmpty(options.appHomeUrl) && ValidUrl(options.appHomeUrl)) {
                this.appHomeUrl = options.appHomeUrl;
            }
            if (!Office.Controls.Utils.isNullOrUndefined(options.customizedItems)) {
                this.customizedItems = options.customizedItems;
            }
            if (!Office.Controls.Utils.isNullOrUndefined(options.onSignIn) && Office.Controls.Utils.isFunction(options.onSignIn)) {
                this.onSignIn = options.onSignIn;
            }
            if (!Office.Controls.Utils.isNullOrUndefined(options.onSignOut) && Office.Controls.Utils.isFunction(options.onSignOut)) {
                this.onSignOut = options.onSignOut;
            }
        }
        if (!Office.Controls.Utils.isNullOrUndefined(loginProvider.hasLogin)) {
            this.isSignedIn = loginProvider.hasLogin();
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
            var instance = this;
            loginProvider.getUserInfoAsync(function(error, userData) {
                if (!Office.Controls.Utils.isNullOrUndefined(userData)) {
                    instance.signedUserInfo = userData;
                    instance.isSignedIn = true;
                } else {
                    instance.isSignedIn = false;
                    Office.Controls.Utils.errorConsole(error);
                }
                instance.updateControl();
            });
        }
    };

    Office.Controls.AppChrome.prototype = {
        rootNode: null,
        dropDownListNode: null,
        loginProvider: null,
        appDisPlayName: "3rd Party App",
        appHomeUrl: "#",
        customizedItems: null,
        isSignedIn: false,
        signedUserInfo: null,
        onSignIn: function() {},
        onSignOut: function() {},
        defaultImage: "data:image/png;base64,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",

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
            this.rootNode.innerHTML = Office.Controls.appChromeTemplates.generateBannerTemplate(this.appDisPlayName, this.appHomeUrl);
            var dropDonwListRoot = document.createElement("div");
            dropDonwListRoot.style.position = 'relative';
            dropDonwListRoot.innerHTML = Office.Controls.appChromeTemplates.generateDropDownList(this.customizedItems);
            this.rootNode.insertBefore(dropDonwListRoot, this.rootNode.childNodes[1]);
            var instance = this;
            document.getElementById('O365_SubLink_ShellSignout').addEventListener('click', function() {
                instance.onSignOut();
                instance.loginProvider.logout();
            });
        },

        updateControl: function() {
            var instance = this;
            var loginButton = document.getElementById('login_user');
            var Personalistview = document.getElementById('_ariaId_7');
            if (this.isSignedIn == false) {
                this.addClass(document.getElementById("o365_ac_loginbutton"), "o365cs-me-tile-nophoto");
                document.getElementById('dropdownIcon').style.display = 'none';
                document.getElementById('user_name').innerText = Office.Controls.Utils.htmlEncode(Office.Controls.appChromeResourceString.SignInString);
                loginButton.addEventListener('click', function() {
                    instance.onSignIn();
                    instance.loginProvider.login();
                });
            } else {
                document.getElementById('user_name').innerText = Office.Controls.Utils.htmlEncode(this.signedUserInfo.displayName);
                if (this.signedUserInfo.imgSrc != null) {
                    instance.defaultImage = this.signedUserInfo.imgSrc;
                }
                document.getElementById('login_user_image').src = this.defaultImage;

                document.getElementById('login_user_image').title = this.signedUserInfo.displayName;
                document.getElementById('image_container').style.display = 'table-cell';
                this.genInlinePersona(document.getElementById('myPersona'));
                loginButton.addEventListener('click', function() {
                    if (Personalistview.style.display == 'none') {
                        Personalistview.style.display = 'block';
                    }else{
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
                this.addClass(document.getElementById('login_user'), "o365cs-personaShow");
                this.removeClass(document.getElementById('login_user'), "ms-bgc-tdr-h");
            } else {
                this.removeClass(document.getElementById('login_user'), "o365cs-personaShow");
                this.addClass(document.getElementById('login_user'), "ms-bgc-tdr-h");
            }
        },

        genInlinePersona: function(ele) {
            if (typeof ele !== 'object') {
                Office.Controls.Utils.errorConsole('Invalid parameters type');
                return;
            }
            var personaObj = {};
            personaObj.jobTitle = null;
            personaObj.department = this.signedUserInfo.accountName;
            personaObj.displayName = this.signedUserInfo.displayName;
            personaObj.imgSrc = this.defaultImage;
            if (this.signedUserInfo.imgSrc != null) {
                personaObj.imgSrc = this.signedUserInfo.imgSrc;
            }
            Office.Controls.Persona.PersonaHelper.createInlinePersona(ele, personaObj);
        },

        hasClass: function(obj, classStr) {
            return obj.className.match(new RegExp('(\\s|^)' + classStr + '(\\s|$)'));
        },

        addClass: function(obj, classStr) {
            if (!this.hasClass(obj, classStr)) obj.className += " " + classStr;
        },

        removeClass: function(obj, classStr) {
            if (this.hasClass(obj, classStr)) {
                var reg = new RegExp('(\\s|^)' + classStr + '(\\s|$)');
                obj.className = obj.className.replace(reg, ' ');
            }
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

        innerHtml += '<div autoid=\"_o365sg2c_1\" class=\"o365cs-me-tileview\"><div class=\"o365cs-me-tileimg\"><img autoid=\"_o365sg2c_5\" class=\"o365cs-me-personaimg\" id=\"login_user_image\"></div></div></div>';

        innerHtml += '<div class=\"o365cs-me-tile-container\"><div autoid=\"_o365sg2c_6\" id=\"o365_ac_loginbutton\"><div class=\"o365cs-me-tile-nophoto-username-container\">';

        innerHtml += '<span autoid=\"_o365sg2c_8\" class=\"o365cs-me-tile-nophoto-username o365cs-me-bidi\" id=\"user_name\"></span></div>';
        innerHtml += '<span class=\"wf-o365-x18 ms-fcl-nt o365cs-me-tile-nophoto-down owaimg wf wf-size-x18\" role=\"presentation\" id=\"dropdownIcon\" style=\"display:table-cell\"><div class=\"o365cs-me-caretDownContainer\"><div class=\"o365cs-me-caretDown\"></div></div></span></div></div>'
        innerHtml += '</button></div></div></div></div></div>';
        return innerHtml;


    };

    Office.Controls.appChromeTemplates.generateDropDownList = function(appLinks) {
        var innerHtml = '<div class=\"o365cs-nav-contextMenu o365cs-dropdownlist contextMenuPopup removeFocusOutline\" ispopup=\"1\" iscontextmenu=\"1\" role=\"menu\" ismodal=\"false\" tabindex=\"-1\" parentids=\"(6)\" style=\"display: none;\" id=\"_ariaId_7\">';
        innerHtml += '<div class=\"o365cs-base ms-bgc-w o365cst o365cs-context-font o365cs-me-contextMenu\"><div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_0\" class=\"o365cs-me-itemsList\" tabindex=\"-1\" id=\"additem\"><div>';
        innerHtml += Office.Controls.appChromeTemplates.generatePersonaPart();
        if (!Office.Controls.Utils.isNullOrUndefined(appLinks)) {
            var hasItem = false;
            for (var name in appLinks) {
                if (isEmpty(name) || isEmpty(appLinks[name]) || !ValidUrl(appLinks[name])) {
                    continue;
                }
                innerHtml += Office.Controls.appChromeTemplates.generateAppLinkPart(name, appLinks[name]);
                hasItem = true;
            }
            if (hasItem) {
                innerHtml += Office.Controls.appChromeTemplates.generateMenuSeparator();
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
        if (name.length >= 40) {
            name = name.substr(0, 40);
        }
        var innerHtml = '<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" class=\"ms-item-tdr\" tabindex=\"-1\" aria-selected=\"false\"><a class=\"o365button o365cs-contextMenuItem ms-fcl-b\" role=\"link\" href=\"' + link + '\">';
        innerHtml += '<div class=\"_fce_j\"><span class=\"_fce_k owaimg\" role=\"presentation\" style=\"display: none;\"></span><span autoid=\"_fce_4\">' + Office.Controls.Utils.htmlEncode(name) + '</span></div></a></div>';
        return innerHtml;
    };

    Office.Controls.appChromeTemplates.generateMenuSeparator = function() {
        var innerHtml = '<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" tabindex=\"-1\"><div class=\"o365cs-contextMenuSeparator ms-bcl-nl\"></div></div>';
        return innerHtml;
    }

    Office.Controls.appChromeTemplates.generateSignOutPart = function() {
        var innerHtml = '<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" class=\"ms-item-tdr\" tabindex=\"-1\" aria-label=\"Sign out and return to the Sign-in page\" title=\"Sign out and return to the Sign-in page\" aria-selected=\"false\"><a class=\"o365button o365cs-contextMenuItem ms-fcl-b\" role=\"link\" id=\"O365_SubLink_ShellSignout\" aria-label=\"Sign out of Office 365 and return to the Sign-in page\">';
        innerHtml += '<div class=\"_fce_j\"><span class=\"_fce_k owaimg\" role=\"presentation\" style=\"display: none;\"></span><span autoid=\"_fce_4\" aria-label=\"Sign out of Office 365 and return to the Sign-in page\">' + Office.Controls.Utils.htmlEncode(Office.Controls.appChromeResourceString.SignOutString) + '</span></div></a></div>';
        return innerHtml;
    };

    Office.Controls.appChromeResourceString = function() {};
    Office.Controls.appChromeResourceString.SignInString = 'Sign in';
    Office.Controls.appChromeResourceString.SignOutString = 'Sign out';


})();
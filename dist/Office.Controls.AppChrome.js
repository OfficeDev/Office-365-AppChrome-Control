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
        defaultImage: "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAGAAAABgCAMAAADVRocKAAAABGdBTUEAALGPC/xhBQAAAAFzUkdCAK7OHOkAAAAgY0hSTQAAeiYAAICEAAD6AAAAgOgAAHUwAADqYAAAOpgAABdwnLpRPAAAAORQTFRFsbGxv7+/wcHBu7u72dnZ7u7u/////Pz85ubmy8vLtbW1wsLC5OTk/f3939/ft7e34eHh+vr61NTUs7OzwMDA8/Pz7OzsxMTE9PT0vLy8+/v7+fn5urq6vb298fHxtLS0srKyzc3N9/f32NjYtra209PT7e3t7+/v/v7+yMjI3t7e4+Pj6+vr6urq4uLi19fXx8fH0tLS9fX1ubm59vb2z8/PuLi4+Pj429vbzs7O0NDQ4ODg1dXV2tra1tbWzMzMxcXF6enp0dHR8vLyysrKxsbG3d3d6Ojovr6+5eXl5+fn8PDwYCkYCwAAAAFiS0dEBmFmuH0AAAAJcEhZcwAAdTAAAHUwAd0zcs0AAAKgSURBVGje7ZjdVtpQEIUByxCIIRKtYECgSSVULbWopRDBarVq+/7vU1raBcTMmbEzueha7Ous/XEOZ35zuY022mij/0z5QqGQmfnWqyL8VskqV/Tt7W1YkVN1de0rO5BQzdtV9N97Dc+1r3eIegnS1DhQ8rd9SFezpeJ/2ARM7Y6Cf7cIuCwFgAcmvRH7t3wjIBCHXBXMCoX+rkMA/LcywBFQOpIBeiQgkt1QnwSAKNre0f6ylxoyAMcSwAkDsC8BnDIAgQQQMQCOBNBmAECSUotZA95zAJIrGjD8mxLABwagKAGcMQADCYCTiz5KAIxsCnkRwCP9hyJ/uqKJayb1UJ1zIeDc3FXAjtCfCoWa9ADzzv3CBLgU+8/HphLuf6rgn8uVUf/GJxUA2hwFOu37rzOk3tJQawCZy075p7eFTeO6OmHiECN5455Q63OwtG+PdWbM+qQaXx0uL2oa96IoGsxWE+h1bB1P7H9x3wpHi59bNQwZ7p+q3fZeOtR+Wan3DXQ/YS+vzYlf8mjd9Z7RuUr9ajdcS+ZOzE5Ml7Xkm7xJidn8s7bMP2PZd+OUqGom01onTKlF/VsOAJn7vpZXnmZrFqR/dUf7j9Hc5lu39UIlfz+Jh/g3ZA/g4psDlhrU5EwNxqSILuCA0WqZRUzOU6k/gPEldUdyQM8E4HS7lPqmtMSZW0mZSsWNBsA0OQdye+P6oqLhDz4O2FMBAJ6273QAdRRAzxssTVDANx3AFAVYOgB88OQsPxh6QAHCWvBX6GyuEwbzooMBHpUAaKTdKwEA61zHcuuFsMlnpgV4zDbO8FzxpAXA2i/Wio4jbHy+kFsv5CEAcrXCFdJ8uVr+WFXOqwFOMg5k+J5xIGMbebVAhh8ZBzK2KVQLZKituP4EqRB824c6sq4AAAAldEVYdGRhdGU6Y3JlYXRlADIwMTQtMTAtMjlUMjA6MjQ6MTktMDU6MDBCpOLkAAAAJXRFWHRkYXRlOm1vZGlmeQAyMDE0LTEwLTI5VDIwOjI0OjE5LTA1OjAwM/laWAAAAABJRU5ErkJggg==",

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
                document.getElementById('login_user_image').style.backgroundImage = "url('" + instance.defaultImage + "')";

                document.getElementById('login_user_image').title = this.signedUserInfo.displayName;
                document.getElementById('image_container').style.display = 'table-cell';
                this.genInlinePersona(document.getElementById('myPersona'));
                loginButton.addEventListener('click', function(e) {
                    if (Personalistview.style.display == 'none') {
                        Personalistview.style.display = 'block';
                    }else{
                        Personalistview.style.display = 'none';
                    }
                    instance.changeTopMenuColor();
                    Office.Controls.Utils.cancelEvent(e);
                });
                document.onclick = function(e) {
                    if (Personalistview.style.display == 'block') {
                        Personalistview.style.display = 'none';
                        instance.changeTopMenuColor();
                    }
                }
                document.oncontextmenu = function (e) {
                    e = e || event;
                    var target = e.target || e.srcElement;
                    if (Personalistview.style.display == 'block') {
                        while (target != undefined && target != null) {
                            if (target == Personalistview) {
                                return true;
                            }
                            if (target == loginButton) {
                                break;
                            }
                            target = target.parentNode;
                        }
                        Personalistview.style.display = 'none';
                        instance.changeTopMenuColor();
                    }
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

        innerHtml += '<div autoid=\"_o365sg2c_1\" class=\"o365cs-me-tileview\"><div class=\"o365cs-me-tileimg\"><div autoid=\"_o365sg2c_5\" class=\"o365cs-me-personaimg\" id=\"login_user_image\"></div></div></div></div>';

        innerHtml += '<div class=\"o365cs-me-tile-container\"><div autoid=\"_o365sg2c_6\" id=\"o365_ac_loginbutton\"><div class=\"o365cs-me-tile-nophoto-username-container\">';

        innerHtml += '<span autoid=\"_o365sg2c_8\" class=\"o365cs-me-tile-nophoto-username o365cs-me-bidi\" id=\"user_name\"></span></div>';
        innerHtml += '<span class=\"wf-o365-x18 ms-fcl-nt o365cs-me-tile-nophoto-down owaimg wf wf-size-x18\" role=\"presentation\" id=\"dropdownIcon\" style=\"display:table-cell\"><div class=\"o365cs-me-caretDownContainer\"><div class=\"o365cs-me-caretDown\"></div></div></span></div></div>'
        innerHtml += '</button></div></div></div></div></div>';
        return innerHtml;


    };

    Office.Controls.appChromeTemplates.generateDropDownList = function(appLinks) {
        var innerHtml = '<div class=\"o365cs-nav-contextMenu o365cs-dropdownlist contextMenuPopup\" ispopup=\"1\" iscontextmenu=\"1\" role=\"menu\" ismodal=\"false\" tabindex=\"-1\" parentids=\"(6)\" style=\"display: none;\" id=\"_ariaId_7\">';
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
        var innerHtml = '<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" class=\"ms-item-tdr\" tabindex=\"-1\" aria-label=\"Sign out\" title=\"Sign out\" aria-selected=\"false\"><a class=\"o365button o365cs-contextMenuItem ms-fcl-b\" role=\"link\" href=\"#\" id=\"O365_SubLink_ShellSignout\" aria-label=\"Sign out of Office 365 and return to the Sign-in page\">';
        innerHtml += '<div class=\"_fce_j\"><span class=\"_fce_k owaimg\" role=\"presentation\" style=\"display: none;\"></span><span autoid=\"_fce_4\" aria-label=\"Sign out of Office 365 and return to the Sign-in page\">' + Office.Controls.Utils.htmlEncode(Office.Controls.appChromeResourceString.SignOutString) + '</span></div></a></div>';
        return innerHtml;
    };

    Office.Controls.appChromeResourceString = function() {};
    Office.Controls.appChromeResourceString.SignInString = 'Sign in';
    Office.Controls.appChromeResourceString.SignOutString = 'Sign out';


})();
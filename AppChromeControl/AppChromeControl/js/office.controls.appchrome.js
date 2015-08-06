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

  Office.Controls.PeopleAadDataProvider = function(authContext) {
    if (Office.Controls.Utils.isFunction(authContext)) {
      this.getTokenAsync = authContext;
    } else {
      this.authContext = authContext;
      if (this.authContext) {
        this.getTokenAsync = function(dataProvider, callback) {
          this.authContext.acquireToken(this.aadGraphResourceId, function(error, token) {
            callback(error, token);
          });
        };
      }
    }
  }

  Office.Controls.PeopleAadDataProvider.prototype = {
    maxResult: 50,
    authContext: null,
    getTokenAsync: undefined,
    aadGraphResourceId: '00000002-0000-0000-c000-000000000000',
    apiVersion: 'api-version=1.5',

    getImageAsync: function(personId, callback) {

      var self = this;
      self.authContext.acquireToken(self.aadGraphResourceId, function(error, token) {

        // Handle ADAL Errors
        if (error || !token) {
          callback('Error', null);
          return;
        }

        var parsed = self.authContext._extractIdToken(token);
        var tenant = '';

        if (parsed) {
          if (parsed.hasOwnProperty('tid')) {
            tenant = parsed.tid;
          }
        }

        var xhr = new XMLHttpRequest();

        xhr.open('GET', 'https://graph.windows.net/' + tenant + '/users/' + personId + '/thumbnailPhoto?' + self.apiVersion);
        xhr.setRequestHeader('Content-Type', 'application/json');
        xhr.setRequestHeader('Authorization', 'Bearer ' + token);
        xhr.responseType = "blob";
        xhr.onabort = xhr.onerror = xhr.ontimeout = function() {
          callback('Error', null);
        };
        xhr.onload = function() {
          if (xhr.status === 401) {
            callback('Unauthorized', null);
            return;
          }
          if (xhr.status !== 200) {
            callback('Unknown error', null);
            return;
          }

          var reader = new FileReader();
          reader.addEventListener("loadend", function() {
            callback(null, reader.result);
          });
          reader.readAsDataURL(xhr.response);
        };
        xhr.send('');
      });
    }
  };

  Office.Controls.appchrome = function(root, login, displayName, ItemtoAdd) {
    if (typeof root !== 'object' || typeof login !== 'object' || typeof displayName !== 'string' || typeof ItemtoAdd !== 'object') {
      Office.Controls.Utils.errorConsole('Invalid parameters type');
      return;
    }
    this.root = root;
    this.loadDorpDownListTemplate(ItemtoAdd);
    this.login = login;
    this.displayName = displayName;
    this.setUserData();
    this.setdisplayName();
  };

  Office.Controls.appchrome.prototype = {
    persona: {
      "ImageUrl": "image/default.jpg",
      "Id": "",
      "PrimaryText": '',
      "SecondaryText": '', // JobTitle, Department
      "SecondaryTextShort": "",
      "TertiaryText": '', // Office
      "Actions": {
        "Email": "",
        "WorkPhone": "",
        "Mobile": "",
        "Skype": ""
      }
    },

    loadDorpDownListTemplate: function(insertElement) {
      try {
        var rootElement = document.getElementById('additem');
        var template = "";
        template = template + Office.Controls.appchrome.Templates.LoginDropDownList["persona"].value;
        template = template + Office.Controls.appchrome.Templates.LoginDropDownList["separator"].value;
        if (!this.isEmptyObject(insertElement)) {
          for (var itemele in insertElement) {
            template = template + this.initItemTemplate(itemele, insertElement[itemele]);
          }
          template = template + Office.Controls.appchrome.Templates.LoginDropDownList["separator"].value;
        }
        template = template + Office.Controls.appchrome.Templates.LoginDropDownList["signout"].value;
        this.parseTemplate(template, rootElement);
      } catch (ex) {
        throw ex;
      }
    },

    parseTemplate: function(templateContent, rootElement) {
      try {
        var templateElement = document.createElement("div");
        templateElement.innerHTML = templateContent;
        if ((Office.Controls.Utils.isNullOrUndefined(templateElement))) {
          Office.Controls.Utils.errorConsole('Fail to get template document');
        }
        rootElement.appendChild(templateElement);
      } catch (ex) {
        throw ex;
      }
    },

    initItemTemplate: function(name, link) {
      var template = "<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" tabindex=\"-1\" aria-selected=\"false\"><a class=\"o365button o365cs-contextMenuItem ms-fcl-b\" role=\"link\" href=\"" + link + "\"><div class=\"_fce_j\"><span class=\"_fce_k owaimg\" role=\"presentation\" style=\"display: none;\"></span><span autoid=\"_fce_4\">" + name + "</span></div></a></div>";
      return template;
    },

    setUserData: function() {
      var user = this.login.getCurrentUser();
      if (user) {
        var appchromeObj = this;
        appchromeObj.persona.SecondaryText = user.userName;
        appchromeObj.persona.PrimaryText = user.profile.family_name + ' ' + user.profile.given_name;
        appchromeObj.persona.Id = user.profile.oid;
        var aadDataProvider = new Office.Controls.PeopleAadDataProvider(appchromeObj.login.authContext);
        aadDataProvider.getImageAsync(user.profile.oid, function(error, imgSrc) {
          if (imgSrc != null) {
            appchromeObj.persona.ImageUrl = imgSrc;
            document.getElementById('login_user_image').src = appchromeObj.persona.ImageUrl;
          }
          document.getElementById('login_user_image').title = appchromeObj.persona.PrimaryText;
        });
        var pcRoot = document.getElementById('myPersona');
        Office.Controls.Persona.PersonaHelper.createInlinePersona(pcRoot, appchromeObj.persona);
        var pcName = document.getElementById('UserName');
        pcName.style.display = "block";
        this.setinnerText(pcName,appchromeObj.persona.PrimaryText);
        //pcName.innerText = appchromeObj.persona.PrimaryText;
      }
    },

    setdisplayName: function() {
      document.getElementById('change_name').innerText = this.displayName;
    },

    setinnerText: function(setroot, textinfo) {
      if (window.navigator.userAgent.toLowerCase().indexOf("firefox") != -1) {
        setroot.innerContent = textinfo;
      } else {
        setroot.innerText = textinfo;
      }
    },

    setloginButton: function() {
      var user = this.login.getCurrentUser();
      var appchromeob = this;
      var loginButton = document.getElementById('login_user');
      var Personalistview = document.getElementById('_ariaId_7');
      loginButton.addEventListener('click', function() {
        if (user) {
          if (Personalistview.style.display == 'none') {
            Personalistview.style.display = 'block';
          } else {
            Personalistview.style.display = 'none';
          }
        } else {
          appchromeob.login.signIn();
        }
      });
      document.getElementById('O365_SubLink_ShellSignout').addEventListener('click', function() {
        if (user) {
          appchromeob.login.signOut();
        }
      });
    },

    setClickEvent: function() {
      var settingsButton = document.getElementById('O365_MainLink_Help');
      var settingslistview = document.getElementById('_ariaId_34');
      var loginButton = document.getElementById('login_user');
      var Personalistview = document.getElementById('_ariaId_7');
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
      }
    },

    isEmptyObject: function(obj) {
      for (var key in obj) {
        return false;
      }
      return true;
    }

  };
})();

Office.Controls.appchrome.createAppchrome = function(root, login, displayName, ItemtoAdd) {
  var appchromeobj = new Office.Controls.appchrome(root, login, displayName, ItemtoAdd);
  //appchromeobj.loadDorpDownListTemplate(ItemtoAdd);
  appchromeobj.setloginButton();
  appchromeobj.setClickEvent();
};
Office.Controls.appchrome.Templates = function() {};
Office.Controls.appchrome.Templates.LoginDropDownList = {
  "persona": {
    value: "<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" tabindex=\"-1\" aria-selected=\"false\"><div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_2\" class=\"o365cs-userInfo o365button\" role=\"group\" tabindex=\"0\"><div class=\"o365cs-me-persona\"><div class=\"o365cs-me-personaView\" id=\"myPersona\"></div></div></div><div class=\"o365button\" role=\"menuitem\" tabindex=\"0\"><div style=\"display: none;\"></div></div><div class=\"o365button o365cs-contextMenuItem ms-fcl-b ms-bgc-nl-h\" role=\"menuitem\" tabindex=\"0\" aria-label=\"Sign in to add another account\" title=\"Sign in to add another account\" style=\"display: none;\"></div><div><div><div class=\"_fce_p ms-bcl-nl\"></div></div></div></div>"
  },
  "separator": {
    value: "<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" tabindex=\"-1\"><div class=\"o365cs-contextMenuSeparator ms-bcl-nl\"></div></div>"
  },
  "signout": {
    value: "<div autoid=\"__Microsoft_O365_ShellG2_Plus_templates_cs_1\" tabindex=\"-1\" aria-label=\"Sign out and return to the Sign-in page\" title=\"Sign out and return to the Sign-in page\" aria-selected=\"false\"><a class=\"o365button o365cs-contextMenuItem ms-fcl-b\" role=\"link\" id=\"O365_SubLink_ShellSignout\" aria-label=\"Sign out of Office 365 and return to the Sign-in page\"><div class=\"_fce_j\"><span class=\"_fce_k owaimg\" role=\"presentation\" style=\"display: none;\"></span><span autoid=\"_fce_4\" aria-label=\"Sign out of Office 365 and return to the Sign-in page\">Sign out</span></div></a></div>"
  }
}
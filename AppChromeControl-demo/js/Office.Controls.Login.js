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

    Office.Controls.ImplicitGrantLogin = function (config) {
        this.authContext = new AuthenticationContext(config);
        this.authContext.handleWindowCallback();
    };

    Office.Controls.ImplicitGrantLogin.prototype = {
        authContext: null,

        signIn: function (callback) {
            var objNull = null;
            if (!Office.Controls.Utils.isNullOrUndefined(callback) && !Office.Controls.Utils.isFunction(callback)) {
                throw new Error('callback is not a function');
            }
            if (this.authContext) {
                if (!Office.Controls.Utils.isNullOrUndefined(callback)) {
                    this.authContext.callback = callback;
                }
                this.authContext.login();
            } else {
                console.log('SignIn failed');
            }
        },

        signOut: function () {
            if (this.authContext && this.authContext.getCachedUser) {
                this.authContext.logOut();
            } else {
                console.log('SignOut failed');
            }
        },

        getAuthContext: function () {
            return this.authContext;
        },

        getAccessTokenAsync: function (resource, callback) {
            if (!Office.Controls.Utils.isFunction(callback)) {
                throw new Error('callback is not a function');
            }
            this.authContext.acquireToken(resource, function (error, token) {
                // Handle ADAL Error
                if (error || !token) {
                    console.log('ADAL Error Occurred: ' + error);
                    return;
                }
                callBack(error, token)
            });
        },

        getCurrentUser: function () {
            return this.authContext.getCachedUser();
        },

        hasSignedIn: function () {
            var user = this.authContext.getCachedUser();
            if (user) {
                return true;
            }else {
                return false;
            }
        },

        getUserInfoAsync: function (callback) {
            var user = this.authContext.getCachedUser()
            var userInfo = new Object();
            userInfo.accountName = user.userName;
            userInfo.displayName = user.profile.family_name + ' ' + user.profile.given_name;
            callback("", userInfo);
    }


    };
})();
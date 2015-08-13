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

    Office.Controls.Login = function(authType, config) {
        this.authType = authType;
        this.authContext = new AuthenticationContext(config);
        this.authContext.handleWindowCallback();
    };

    Office.Controls.Login.prototype = {
        authType: "implicit",
        authContext: null,
        aadGraphResourceId: '00000002-0000-0000-c000-000000000000',
        apiVersion: 'api-version=1.5',

        signIn: function(callback) {
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

        signOut: function() {
            if (this.authContext && this.authContext.getCachedUser) {
                this.authContext.logOut();
            } else {
                console.log('SignOut failed');
            }
        },

        getAuthContext: function() {
            return this.authContext;
        },

        getAccessTokenAsync: function(resource, callback) {
            if (!Office.Controls.Utils.isFunction(callback)) {
                throw new Error('callback is not a function');
            }
            this.authContext.acquireToken(resource, function(error, token) {
                // Handle ADAL Error
                if (error || !token) {
                    console.log('ADAL Error Occurred: ' + error);
                    return;
                }
                callBack(error, token)
            });
        },

        getCurrentUser: function() {
            return this.authContext.getCachedUser();
        },

        hasSignedIn: function() {
            var user = this.authContext.getCachedUser();
            if (user) {
                return true;
            } else {
                return false;
            }
        },

        getUserImageAsync: function(personId, callback) {
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
        },

        getUserInfoAsync: function(callback) {
            var user = this.authContext.getCachedUser()
            var userInfo = new Object();
            userInfo.accountName = user.userName;
            userInfo.displayName = user.profile.family_name + ' ' + user.profile.given_name;
            this.getUserImageAsync(user.profile.oid, function(error, image) {
                userInfo.imgSrc = image;
                callback(error, userInfo);
            });
        }

    };
})();
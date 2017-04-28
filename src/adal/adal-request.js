var q = require('q');
var AuthenticationContext = require('adal-angular');
var adalConfig = require('./adal-config');

var _adal = new AuthenticationContext(adalConfig);


/*
OAuth flows for authentication and retrieving access tokens for the different resources are based on callbacks. 
The web application redirects to Azure Active Directory which executes the specific flow and redirects back to the application 
including information specific to the particular flow. 
For an application to use OAuth it has to handle the callbacks and process the information retrieved from AAD.

The processAdalCallback function executes only if an AAD OAuth hash has been provided in the URL (line 11). 
This is how AAD passes its information to web applications in the implicit OAuth flow and this is what the function is checking for. 
Based on the provided hash ADAL JS tries to determine the type of flow that the hash belongs to (line 13). 
After validating the hash (line 26) any callbacks registered with ADAL JS will be executed completing the previously started flow. 
*/
var processAdalCallback = function () {
    var hash = window.location.hash;

    if (_adal.isCallback(hash)) {
        // callback can come from login or iframe request
        var requestInfo = _adal.getRequestInfo(hash);
        _adal.saveTokenFromHash(requestInfo);
        window.location.hash = '';

        if (requestInfo.requestType !== _adal.REQUEST_TYPE.LOGIN) {
            if (window.parent.AuthenticationContext === 'function' && window.parent.AuthenticationContext()) {
                _adal.callback = window.parent.AuthenticationContext().callback;
            }
            if (requestInfo.requestType === _adal.REQUEST_TYPE.RENEW_TOKEN) {
                _adal.callback = window.parent.callBackMappedToRenewStates[requestInfo.stateResponse];
            }
        }

        if (requestInfo.stateMatch) {
            if (typeof _adal.callback === 'function') {
                // Call within the same context without full page redirect keeps the callback
                if (requestInfo.requestType === _adal.REQUEST_TYPE.RENEW_TOKEN) {
                    // Idtoken or Accestoken can be renewed
                    if (requestInfo.parameters['access_token']) {
                        _adal.callback(_adal._getItem(_adal.CONSTANTS.STORAGE.ERROR_DESCRIPTION), requestInfo.parameters['access_token']);
                        return;
                    } else if (requestInfo.parameters['id_token']) {
                        _adal.callback(_adal._getItem(_adal.CONSTANTS.STORAGE.ERROR_DESCRIPTION), requestInfo.parameters['id_token']);
                        return;
                    }
                }
            } else {
                // normal full login redirect happened on the page
                updateDataFromCache(_adal.config.loginResource);
                if (_oauthData.userName) {
                    //IDtoken is added as token for the app
                    window.setTimeout(function () {
                        updateDataFromCache(_adal.config.loginResource);
                        // redirect to login requested page
                        var loginStartPage = _adal._getItem(_adal.CONSTANTS.STORAGE.START_PAGE);
                        if (loginStartPage) {
                            window.location.path(loginStartPage);
                        }
                    }, 1);
                }
            }
        }
    }
}

// The isAuthenticated function returns a promise that resolves whenever the user gets authenticated.
var isAuthenticated = function () {
    var deferred = q.defer();

    updateDataFromCache(_adal.config.loginResource);
    if (!_adal._renewActive && !_oauthData.isAuthenticated && !_oauthData.userName) {
        if (!_adal._getItem(_adal.CONSTANTS.STORAGE.FAILED_RENEW)) {
            // Idtoken is expired or not present
            _adal.acquireToken(_adal.config.loginResource, function (error, tokenOut) {
                if (error) {
                    _adal.error('adal:loginFailure', 'auto renew failure');
                    deferred.reject();
                } else {
                    if (tokenOut) {
                        _oauthData.isAuthenticated = true;
                        deferred.resolve();
                    } else {
                        deferred.reject();
                    }
                }
            });
        } else {
            deferred.resolve();
        }
    } else {
        deferred.resolve();
    }

    return deferred.promise;
}

// This function is responsible for executing a web request to and endpoint secured with AAD such as 
// SharePoint Online API or Microsoft Graph. Such requests can only be executed by users who are authenticated 
// and who have a valid access token to the specific endpoint. 
var adalRequest = function (settings) {
    var deferred = q.defer();

    isAuthenticated().then(function () {
        var resource = _adal.getResourceForEndpoint(settings.url);

        if (!resource) {
            _adal.info('No resource configured for \'' + settings.url + '\'');
            deferred.reject();
            return deferred.promise;
        }

        var tokenStored = _adal.getCachedToken(resource);
        if (tokenStored) {
            if (!settings.headers) {
                settings.headers = {};
            }

            settings.headers.Authorization = 'Bearer ' + tokenStored;

            makeRequest(settings).then(deferred.resolve, deferred.reject);
        } else {
            var isEndpoint = false;

            for (var endpointUrl in _adal.config.endpoints) {
                if (settings.url.indexOf(endpointUrl) > -1) {
                    isEndpoint = true;
                }
            }

            if (_adal.loginInProgress()) {
                _adal.info('Login already in progress');
                deferred.reject();
            } else if (isEndpoint) {
                _adal.acquireToken(resource, function (error, tokenOut) {
                    if (error) {
                        deferred.reject();
                        _adal.error(error);
                    } else {
                        if (tokenOut) {
                            _adal.verbose('Token is available');
                            if (!settings.headers) {
                                settings.headers = {};
                            }
                            settings.headers.Authorization = 'Bearer ' + tokenOut;
                            makeRequest(settings).then(deferred.resolve, deferred.reject);
                        }
                    }
                });
            }
        }
    }, function () {
        _adal.login();
    })

    return deferred.promise;
}

module.exports = {
    adalRequest: adalRequest,
    processAdalCallback: processAdalCallback
}
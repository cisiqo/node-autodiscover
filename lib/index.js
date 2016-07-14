var _ = require('underscore');
var validator = require('validator');
var request = require('request');
var async = require('async');
var xmlbuilder = require('xmlbuilder');
var parseString = require('xml2js').parseString;

var getEwsUrl = function(mailAddress, auth, callback) {
  if (!mailAddress) {
    callback(new Error('mailAddress is needed'));
    return;
  }
  if (!auth || !auth.username) {
    callback(new Error('username is needed'));
    return;
  }
  if (!auth || !auth.password) {
    callback(new Error('password is needed'));
    return;
  }
  if (!validator.isEmail(mailAddress)) {
    callback(new Error('Invalid format: ' + mailAddress));
    return;
  }

  var autodiscover = createAutodiscoverService(mailAddress, auth.username, auth.password);
  autodiscover.trySecureUrls(function(err, url) {
    if (err) {
      autodiscover.tryRedirectionUrls(function(err, url) {
        if (err) {
          callback(err);
          return;
        }
        callback(null, url);
      });
      return;
    }
    callback(null, url);
  });
};

var createAutodiscoverService = function(mailAddress, username, password) {
  var smtpDomain = pickSmtpDomain(mailAddress);
  var trySecureUrls = function(callback) {
    async.waterfall([
      function(cb) {
        sendRequest('https://' + smtpDomain + '/autodiscover/autodiscover.xml', function(err, url) {
          if (err) {
            cb(null, null);
            return;
          }
          cb(null, url);
        });
      },
      function(url, cb) {
        if (url) {
          cb(null, url);
          return;
        }

        sendRequest('https://autodiscover.' + smtpDomain + '/autodiscover/autodiscover.xml', function(err, newurl) {
          if (err) {
            cb(err);
            return;
          }
          cb(null, newurl);
        });
      }
    ], function(err, resultUrl) {
      if (err) {
        callback(err);
        return;
      }
      callback(null, resultUrl);
    });
  };

  var tryRedirectionUrls = function(callback) {
    request({
      url: 'http://autodiscover.' + smtpDomain + '/autodiscover/autodiscover.xml',
      method: 'GET',
      followRedirect: false,
      timeout: 10000
    }, function(err, res, body) {
      if (err) {
        callback(err);
        return;
      }
      if (res.statusCode !== 302) {
        callback(new Error('Not redirect'));
        return;
      }
      sendRequest(res.headers.location, function(err, url) {
        if (err) {
          callback(err);
          return;
        }
        callback(null, url);
      });
    });
  };

  var sendRequest = function(url, callback) {
    request({
      url: url,
      method: 'POST',
      headers: {
        'Content-Type': 'text/xml; charset=utf-8'
      },
      body: buildRequestBody(mailAddress),
      auth: {
        username: username,
        password: password
      },
      timeout: 10000
    }, function(err, res, body) {
      if (err) {
        callback(err);
        return;
      }
      if (res.statusCode === 302) {
        sendRequest(res.headers.location, callback);
        return;
      }
      parseRequestBody(body, function(err, result) {
        if (err) {
          callback(err);
          return;
        }

        var error = pickErrorResult(result);
        if (error) {
          callback(new Error('errorCode: ' + error.errorCode + ', message: "' + error.message + '"'));
          return;
        }

        var ewsUrl;
        try {
          ewsUrl = pickEwsUrl(result);
        } catch (e) {
          callback(new Error('Failed to get EwsUrl'));
          return;
        }
        callback(null, ewsUrl);
      });
    });
  };

  var buildRequestBody = function(mailAddress) {
    var root = xmlbuilder.create('Autodiscover').att(
      'xmlns', 'http://schemas.microsoft.com/exchange/autodiscover/mobilesync/requestschema/2006');
    var request = root.ele('Request');
    request.ele('EMailAddress', {}, mailAddress);
    request.ele(
      'AcceptableResponseSchema',
      {},
      'http://schemas.microsoft.com/exchange/autodiscover/mobilesync/responseschema/2006');
    return root.end({pretty: true});
  };

  var parseRequestBody = function(body, callback) {
    parseString(body, function(err, result) {
      if (err) {
        callback(err);
        return;
      }
      callback(null, result);
    });
  };

  var pickErrorResult = function(result) {
    var errors = result.Autodiscover.Response[0].Error;
    if (!errors) {
      return null;
    }
    return {
      errorCode: errors[0].ErrorCode[0],
      message: errors[0].Message[0]
    };
  };

  var pickEwsUrl = function(result) {
    return result.Autodiscover.Response[0].Action[0].Settings[0].Server[0].Url[0];
  };

  return {
    trySecureUrls: trySecureUrls,
    tryRedirectionUrls: tryRedirectionUrls
  };
};

var pickSmtpDomain = function(mailAddress) {
  return _.last(mailAddress.split('@'));
};

exports.getEwsUrl = getEwsUrl;

// Copyrights belong to Discover Technologies LLC. © 2019 Discover Technologies LLC
// Script Continued Below
var DocIntADAL_DS = (function () {
    var _instance = null;

    DocIntADAL_DS = function () {
        return DocIntADAL_DS.prototype._instance;
    };

    DocIntADAL_DS = function (clientId, tenant, spUri) {
        this.config = {
            'tenant': tenant,
            'clientId': clientId,
            'redirectUri': window.location.origin + '/',
            //'redirectUri': 'https://dev70860.service-now.com/sp?id=docintegrator_example_1&table=incident&sys_id=9c573169c611228700193229fff72400&view=sp&defaultTabs',
            'postLogoutRedirectUri': window.location.origin + '/',
            'endpoints': {
                'graphApiUri': 'https://graph.microsoft.com',
                'sharePointUri': spUri
            },
            'cacheLocation': 'localStorage',
            popUp: false,
            //Custom PopUP Method
           // displayCall: function (urlNavigate) {
                //var loginUrl = urlNavigate;
                //var actx = this;
                //$sce.trustAsResourceUrl(loginUrl);
                //$scope.url = loginUrl;
                //var popupWindow = window.open(loginUrl, "login", 'width=483, height=600');

                //per cert.
                //popupWindow.opener = null;
                //roll back
                if (popupWindow && popupWindow.focus)
                    popupWindow.focus();
                var registeredRedirectUri = this.redirectUri;

                var pollTimer = window.setInterval(function () {
                    if (!popupWindow || popupWindow.closed || popupWindow.closed === undefined) {
                        window.clearInterval(pollTimer);
                        authWait = false;
                    }

                    try {
                        if (popupWindow.document.URL.indexOf(registeredRedirectUri) != -1) {
                            window.clearInterval(pollTimer);

                            var dta = new DocIntADAL_DS();

                            if (dta.authContext.isCallback(popupWindow.location.hash)) {
                                var reqInfo = dta.authContext.getRequestInfo(popupWindow.location.hash);



                                dta.authContext.saveTokenFromHash(reqInfo);

                                //window.location = authContext._getItem(authContext.CONSTANTS.STORAGE.LOGIN_REQUEST);
                                //authContext.login();
                                //authContext.saveTokenFromHash(popupWindow.location.hash);
                            }

                            var userctx = dta.authContext.getCachedUser();
                            window.location.hash = popupWindow.location.hash;

                            window.postMessage(popupWindow.location.hash, window.location.origin);

                            popupWindow.close();
                        }
                    } catch (e) {
                        //alert('error $$$ ' + e);
                    }
                }, 20);
            }
        };

        if (DocIntADAL_DS.prototype._instance) {
            return DocIntADAL_DS.prototype._instance;
        }

        this.authContext = new AuthenticationContext(this.config);

        DocIntADAL_DS.prototype._instance = this;

        return DocIntADAL_DS.prototype._instance;
    };

    DocIntADAL_DS.prototype.context = new function () {
        return DocIntADAL_DS.prototype._instance;
    };

    DocIntADAL_DS.prototype.log = function (msg) { };

    DocIntADAL_DS.prototype.validateUser = function () {
        var user = this.authContext.getCachedUser();

        // no user found
        if (user == null) {
            this.authContext.login();

            // todo: come up with a cleaner auth check
            user = this.authContext.getCachedUser();
        }

        return user;
    };

    DocIntADAL_DS.prototype.getToken = function (callback) {
        var user = this.validateUser();
        // var cachedToken = this.adal.getCachedToken(client_id_goes_here);
        // if (cachedToken) {
        //    this.adal.acquireToken("https://graph.microsoft.com", function(error, token) {
        //         jslog(error);
        //         jslog(token);
        //     });
        // }



        // no user found
        if (user != null) {
            //var cachedToken = this.authContext.getCachedToken (this.authContext.config.clientId);
            // ensure we have made the token request
            this.authContext.acquireToken(this.authContext.config.endpoints.sharePointUri, function (error, token) {
                // this.authContext.acquireToken(this.authContext.config.clientId, function (error, token) {
                if (error) {

                    // todo: ref a singletone instance
                    var authContext = new AuthenticationContext();

                    // cheesy
                    if (error == 'Token renewal operation failed due to timeout') {
                        //authContext.acquireTokenPopup(authContext.config.clientId, null, null, function (error, token) {
                        authContext.acquireTokenPopup(authContext.config.endpoints.sharePointUri, null, null, function (error, token) {
                            if (error) {
                                //alert('error === ' + error);
                            } else {
                                // authorization
                                var bearer = 'Bearer ' + token;

                                if (callback != null) {
                                    callback(bearer);
                                }
                            }
                        });
                    }
                } else {
                    // authorization
                    var bearer = 'Bearer ' + token;

                    if (callback != null) {
                        callback(bearer);
                    }
                }
            });
        }

        return user;
    };

    return DocIntADAL_DS;
}());


/* JSONPath 0.8.0 - XPath for JSON
     *
     * Copyright (c) 2007 Stefan Goessner (goessner.net)
     * Licensed under the MIT (MIT-LICENSE.txt) licence.
     */
function jsonPath(obj, expr, arg) {
    var P = {
        resultType: arg && arg.resultType || "VALUE",
        result: [],
        normalize: function (expr) {
            var subx = [];
            return expr.replace(/[\['](\??\(.*?\))[\]']/g, function ($0, $1) { return "[#" + (subx.push($1) - 1) + "]"; })
                .replace(/'?\.'?|\['?/g, ";")
                .replace(/;;;|;;/g, ";..;")
                .replace(/;$|'?\]|'$/g, "")
                .replace(/#([0-9]+)/g, function ($0, $1) { return subx[$1]; });
        },
        asPath: function (path) {
            var x = path.split(";"), p = "$";
            for (var i = 1, n = x.length; i < n; i++)
                p += /^[0-9*]+$/.test(x[i]) ? ("[" + x[i] + "]") : ("['" + x[i] + "']");
            return p;
        },
        store: function (p, v) {
            if (p) P.result[P.result.length] = P.resultType == "PATH" ? P.asPath(p) : v;
            return !!p;
        },
        trace: function (expr, val, path) {
            if (expr) {
                var x = expr.split(";"), loc = x.shift();
                x = x.join(";");
                if (val && val.hasOwnProperty(loc))
                    P.trace(x, val[loc], path + ";" + loc);
                else if (loc === "*")
                    P.walk(loc, x, val, path, function (m, l, x, v, p) { P.trace(m + ";" + x, v, p); });
                else if (loc === "..") {
                    P.trace(x, val, path);
                    P.walk(loc, x, val, path, function (m, l, x, v, p) { typeof v[m] === "object" && P.trace("..;" + x, v[m], p + ";" + m); });
                }
                else if (/,/.test(loc)) { // [name1,name2,...]
                    for (var s = loc.split(/'?,'?/), i = 0, n = s.length; i < n; i++)
                        P.trace(s[i] + ";" + x, val, path);
                }
                else if (/^\(.*?\)$/.test(loc)) // [(expr)]
                    P.trace(P.eval(loc, val, path.substr(path.lastIndexOf(";") + 1)) + ";" + x, val, path);
                else if (/^\?\(.*?\)$/.test(loc)) // [?(expr)]
                    P.walk(loc, x, val, path, function (m, l, x, v, p) { if (P.eval(l.replace(/^\?\((.*?)\)$/, "$1"), v[m], m)) P.trace(m + ";" + x, v, p); });
                else if (/^(-?[0-9]*):(-?[0-9]*):?([0-9]*)$/.test(loc)) // [start:end:step]  phyton slice syntax
                    P.slice(loc, x, val, path);
            }
            else
                P.store(path, val);
        },
        walk: function (loc, expr, val, path, f) {
            if (val instanceof Array) {
                for (var i = 0, n = val.length; i < n; i++)
                    if (i in val)
                        f(i, loc, expr, val, path);
            }
            else if (typeof val === "object") {
                for (var m in val)
                    if (val.hasOwnProperty(m))
                        f(m, loc, expr, val, path);
            }
        },
        slice: function (loc, expr, val, path) {
            if (val instanceof Array) {
                var len = val.length, start = 0, end = len, step = 1;
                loc.replace(/^(-?[0-9]*):(-?[0-9]*):?(-?[0-9]*)$/g, function ($0, $1, $2, $3) { start = parseInt($1 || start); end = parseInt($2 || end); step = parseInt($3 || step); });
                start = (start < 0) ? Math.max(0, start + len) : Math.min(len, start);
                end = (end < 0) ? Math.max(0, end + len) : Math.min(len, end);
                for (var i = start; i < end; i += step)
                    P.trace(i + ";" + expr, val, path);
            }
        },
        eval: function (x, _v, _vname) {
            try { return $ && _v && eval(x.replace(/@/g, "_v")); }
            catch (e) { throw new SyntaxError("jsonPath: " + e.message + ": " + x.replace(/@/g, "_v").replace(/\^/g, "_a")); }
        }
    };

    var $ = obj;
    if (expr && obj && (P.resultType == "VALUE" || P.resultType == "PATH")) {
        P.trace(P.normalize(expr).replace(/^\$;/, ""), obj, "$");
        return P.result.length ? P.result : false;
    }
}

function validateLibrary(g_form) {
    var isValid = false;

    var currentLibraryId = g_form.getUniqueValue();

    //sys_id of connection
    var connection = g_form.getValue('connectionid');
    var conn_base_url = g_scratchpad.conn_base_url;

    var relative_url = g_form.getValue('relative_url');
    var list_title = g_form.getValue('list_title');

    var one_drive_id = g_form.getValue('one_drive_id');
    var content_type = g_form.getValue('content_type');
    var sp_site_id = g_form.getValue('sp_site_id');
    var sp_library_id = g_form.getValue('library_id');

    var connectionType = g_scratchpad.conn_type;

    debugger;

    if (connectionType != 'o365') {

        //alert('Test Library in progress for ' + connectionType);

        var sp = new DocIntSharePoint(conn_base_url, relative_url, list_title, connectionType,sp_library_id);

        var result = sp.testLibraryNoJQuery(null);

        if (result && !result.error && result.data) {

            var responseDataObj = JSON.parse(result.data);

            if (responseDataObj.d) {
                g_form.addInfoMessage(DOCINT_CONSTANT_MSG_TEST_LIB_OK);
                isValid = true;

                var splibraryId = responseDataObj.d.Id;
                var itemType = responseDataObj.d.ListItemEntityTypeFullName;

                //         libaryResult.sp_library_id = resultD.Id;
                //         libaryResult.sp_item_type = resultD.ListItemEntityTypeFullName;
                //update additional meta-data from test library call
                var ga = new GlideAjax('ValidateConnectionAjaxService');
                ga.addParam('sysparm_name', 'updateSPLibrary');

                ga.addParam('sysparm_currentLibrarySysId', currentLibraryId);
                ga.addParam('sysparm_sp_library_id', splibraryId);
                ga.addParam('sysparm_sp_item_type', itemType);

                ga.getXML(function () { });
            }
            else {
                g_form.addErrorMessage("Failed retrieving library list from SharePoint response \n" + result.data);

            }
        }
        else {
            g_form.addErrorMessage(result.error);
        }
    } else if (connectionType == 'o365') {
        //alert('Test Library to SharePoint Online');

        var ga = new GlideAjax('ValidateConnectionAjaxService');
        ga.addParam('sysparm_name', 'validateSPOLibrary');

        ga.addParam('sysparm_currentConnectionId', connection);
        ga.addParam('sysparm_relative_url', relative_url);
        ga.addParam('sysparm_list_title', list_title);
        ga.addParam('sysparm_content_type', content_type);
        ga.addParam('sysparm_one_drive_id', one_drive_id);
        ga.addParam('sysparm_sp_site_id', sp_site_id);
        ga.addParam('sysparm_sp_library_id', sp_library_id);

        ga.getXMLWait();

        //ga.getXML(processAjaxResponse);

        var answer = ga.getAnswer();

        if (answer) {
            isValid = true;
            g_form.addInfoMessage(answer);
        }
        else {
        }

        //async
        return isValid;
    }

    // function processAjaxResponse(validationResponse) {
    //     var answer;
    //     if (validationResponse && validationResponse.responseXML) {
    //         answer = validationResponse.responseXML.documentElement.getAttribute("answer");
    //     } else {
    //         answer = validationResponse;
    //     }

    //     isValid = true;
    //     g_form.addInfoMessage(answer);

    //     //async
    //     return isValid;
    // }

}

function getContentType(g_form, expr, arg) {
    var currentLibraryId = g_form.getUniqueValue();

    //sys_id of connection
    var connection = g_form.getValue('connectionid');
    var conn_base_url = g_scratchpad.conn_base_url;

    var library_name = g_form.getValue('name');
    var relative_url = g_form.getValue('relative_url');
    var list_title = g_form.getValue('list_title');

    var one_drive_id = g_form.getValue('one_drive_url');
    var content_type = g_form.getValue('content_type');
    var sp_site_id = g_form.getValue('sp_site_id');
    var sp_library_id = g_form.getValue('library_id');

    var connectionType = g_scratchpad.conn_type;

    var client_id = g_scratchpad.conn_client_id;
    var tenant_name = g_scratchpad.conn_tenant_name;

    var caller = g_form.getReference("connectionid", doProcess);

    debugger;

    function doProcess(caller) {
        //alert(" curent content_type  " + content_type);

        connectionType = caller.conn_type;
        var contentTypes = null;

        if (connectionType != 'o365') {
            //alert('Test Library to SP2013 in progress');
            var sp = new DocIntSharePoint(conn_base_url, relative_url, list_title, connectionType, sp_library_id);

            contentTypes = sp.getSPContentTypeNoJQuery(null);
        } else {
            var _dtadlal = new DocIntADAL_DS(client_id, tenant_name, conn_base_url);

            var dtsp = new DocIntSharePoint(conn_base_url, relative_url, list_title, connectionType, sp_library_id);

            _dtadlal.getToken(function (token) {

                contentTypes = dtsp.getSPContentTypeNoJQuery(token);

            }, _dtadlal.authContext);
        }

        var spContentTypeResponseObject = JSON.parse(contentTypes);

        var contentTypesBlob = [];

        var currValue = g_form.getValue('content_type');
        g_form.clearOptions('content_type');

        g_form.addOption('content_type', 'ALL', 'ALL');
        //if (spContentTypeResponseObject) g_form.clearOptions('content_type');

        for (var j = 0; j < spContentTypeResponseObject.d.results.length; j++) {

            var curr = spContentTypeResponseObject.d.results[j];

            var contentTypeElement = {};

            //Mapping Rules
            contentTypeElement.Name = curr.Name;
            contentTypeElement.ContentTypeId = curr.Id.StringValue;
            contentTypeElement.Selected = false;

            contentTypesBlob.push(contentTypeElement);

            //alway set
            g_form.addOption('content_type', contentTypeElement.Name, contentTypeElement.Name);

            if (currValue != null) {
                if (currValue != contentTypeElement.Name) {
                    //g_form.addOption('content_type', contentTypeElement.Name, contentTypeElement.Name);
                }
                else {
                    // g_form.addOption('content_type', contentTypeElement.Name, contentTypeElement.Name);
                    g_form.setValue('content_type', currValue);
                }
            }
            else {
                //null, never set before
                //g_form.addOption('content_type', contentTypeElement.Name, contentTypeElement.Name);
            }


        }

        var ga = new GlideAjax('ContentTypeAjaxService');
        ga.addParam('sysparm_name', 'insertContentTypes');

        ga.addParam('sysparm_libSysId', currentLibraryId);
        ga.addParam('sysparm_content_type_blob', JSON.stringify(contentTypesBlob, null, 2));

        ga.getXML(processGetContentTypeAjaxResponse);

    }
}

function processGetContentTypeAjaxResponse(validationResponse) {
    var answer = validationResponse.responseXML.documentElement.getAttribute("answer");
    g_form.clearMessages();
    g_form.addInfoMessage(" " + answer);
}

function populateDropdownFromJSONBlob(g_form, blob_field, dropdown_field, json_field) {
    // var lt_blob = g_form.getValue('sp_list_blob');

    // if (lt_blob) {
    //    var jsonLtBlob = JSON.parse(lt_blob);

    //    for (var key in jsonLtBlob) {
    //       if (jsonLtBlob.hasOwnProperty(key)) {
    //          var obj = jsonLtBlob[key];
    //          g_form.addOption('list_title', obj.Name, obj.Name);
    //       }
    //    }
    // }
    var blob = g_form.getValue(blob_field);

    if (blob) {
        var jsonBlob = JSON.parse(blob);

        for (var key in jsonBlob) {
            if (jsonBlob.hasOwnProperty(key)) {
                var obj = jsonBlob[key];

                var f = "Name";

                if (json_field) {
                    f = json_field;
                }

                g_form.addOption(dropdown_field, obj[f], obj[f]);
            }
        }
    }

}

//called by UI Macro, CommonMacro
function renderIframe(anchor, document) {

    debugger;
    // do not remove this empty script,  it is used to find the formatter.
    //var anchor = "${jvar_var_anchor}";

    var var_anchor = $j("#" + anchor);

    var closeSpan = var_anchor.closest("span")[0];
    var tabName = closeSpan.children[0].children[0].innerText;
    var sectionId = closeSpan.id;
    var sectionSysId = sectionId.substring(8);

    //alert('DocInt Common Macro, Section ' + tabName + " section id " + sectionSysId );

    var ga = new GlideAjax('PlatIntAjaxService');
    ga.addParam('sysparm_name', 'getSectionRegistrationInfo');
    ga.addParam('sysparm_form_section_sysid', sectionSysId);

    ga.getXMLWait();
    var configItems = ga.getAnswer();

    //alert(configItems);
    var itemObj = JSON.parse(configItems);

    // <!-- platIntItem.name = ga.getDisplayValue('display_name');
    //    platIntItem.widget_name = ga.getValue('docintegrator_widget');
    //     platIntItem.view_id = ga.getValue('docintegrator_view');
    //     platIntItem.active = ga.getValue('active'); -->

    var platIntItem = itemObj[0];
    //alert(platIntItem.widget_name);

    var isActive = (platIntItem.active == 1);

    var target = anchor + "_iframe1";

    var widget = platIntItem.widget_name;
    var view_id = platIntItem.view_id;
    var page_name = platIntItem.page;

    var timeout_time = 1000;

    var parentURL;
    var table;
    var sys_id;

    parentURL = location.href;

    var url = new URL(parentURL);
    sys_id = url.searchParams.get('sys_id') || '';
    table = url.searchParams.get('sysparm_record_target') || '';

    var defaultPortal = "docint_pi_portal";

    if (widget == "list_view_widget") {
        var map_sys_id;

        var timeout = setTimeout(function () {

            var iframeElement = document.getElementById(target);

            //iframeElement.src="/rc?id=listview_integration&amp;sysparm_domain_restore=false&amp;sysparm_stack=no";
            
            var page = "listview_integration";
            
            if(page_name) {
                page = page_name;
            }
            
            //var src_prefix = "/discoverdocintegratorportal?id=" + page + "&native=true&sysparm_domain_restore=false&sysparm_stack=no";
            var src_prefix = "/" + defaultPortal + "?id=" + page + "&native=true&sysparm_domain_restore=false&sysparm_stack=no";

            if (isActive) {
                iframeElement.src = src_prefix + "&table=" + table + "&map_sys_id=" + sys_id + "&vw=" + view_id;
            }
            else {
                iframeElement.style.display = "none";
            }

        }, timeout_time);
    }
    else if (widget == "search_page_widget") {
        var recTarget;
        var query;
        var sys_id;

        var search_source = platIntItem.search_source;
        var search_field = platIntItem.search_field;

        //Entering  Search Macro
        var interval = setInterval(function () {
            if (document.readyState === 'complete') {
                debugger;
                clearInterval(interval);

                var iframeElement = document.getElementById(target);

                var page = "search_integration";
            
                if(page_name) {
                    page = page_name;
                }

                //var src_prefix = "/discoverdocintegratorportal?id=search_native&native=true&t=" + search_source + "&q=";
                var src_prefix = "/" + defaultPortal + "?id=" +  page + "&native=true&t=" + search_source + "&q=";

                if (isActive) {

                    //query = document.getElementById(table + '.short_description').value;
                    //

                    query = document.getElementById(table + '.' + search_field).value;

                    console.log("query: " + query);

                    var appendStr = query + "&sys_id=" + sys_id + "&vw=" + view_id;;

                    iframeElement.src = src_prefix + appendStr;
                }
                else {
                    iframeElement.style.display = "none";
                }
            }
        }, 1000);
    }
}

const DOCINT_CONSTANT_PREV_KEY = "sn.library.form.list.title.prev";
const DOCINT_CONSTANT_MSG_TEST_LIB_OK = "Test Library Succeeds!";

//supported file format
//https://docs.microsoft.com/en-us/officeonlineserver/office-online-server-overview
const OWA_SUPPORTED_FILE_FORMATS = "pdf,doc,docx,dotx,dot,dotm,xls,xlsx,xlsm,xlm,xlsb,ppt,pptx,pps,ppsx,potx,pot,pptm,potm,ppsm";


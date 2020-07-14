// Copyrights belong to Discover Technologies LLC. © 2019 Discover Technologies LLC
// Script Continued Below
//----------------------------------------------------------------------
// DocIntSharePoint v1.0.0
// @preserve Copyright (c) Discover Technologies, LLC
// @author Discover Technologies - Product Development
//----------------------------------------------------------------------

var DocIntSharePoint = (function () {
    //'use strict';

    /**
     * @constructor
     * @param {string} refSite - the root web URL of the SharePoint site
     * @param {string} refList = the Document Library name related to the Tasker connection
     */
    DocIntSharePoint = function (refBase, refSite, refListPath, spVersion, refListID) {
        this.baseUrl = refBase; //https://host
        this.site = refSite; // '/sites/ServiceNow			
        this.listPath = refListPath; //'lib21'
        //this.list refer to listPath
        this.listId = refListID;

        this.sharePointVersion = spVersion; //o365, sp2013, sp2016,
    };

    /**
     * builds a REST uri using the site and specified endpoint
     * @param {any} endpooint
     */
    DocIntSharePoint.prototype.buildSPRestUri = function (endpoint) {
        var url = this.baseUrl + this.site;
        if (!url) {
            url = "";
        }

        // remove the trailing /if it exists
        if (url[url.length - 1] == '/') {
            url = url.substring(0, url.length - 1);
        }

        // add the endpoint
        url += endpoint;

        // return the result
        return url;
    };

    DocIntSharePoint.prototype.buildAuthToken = function (token) {
        var bearer = token;

        // ensure the token contains the Bearer prefix
        if (!bearer.startsWith('Bearer')) {
            bearer = 'Bearer ' + bearer;
        }
        // return the auth token
        return bearer;
    };

    DocIntSharePoint.prototype.getThisList = function () {
        return "/_api/web/Lists(guid'" + this.listId + "')";
    };

    DocIntSharePoint.prototype.getCORSAuthHeaders = function (digestToken) {
        var headers1 = {
            "Content-Type": "application/json;odata=verbose",
            "Accept": "application/json;odata=verbose",
            "crossDomain": "true",
            "credentials": "include"
        };

        if (digestToken) {
            headers1["X-RequestDigest"] = digestToken;
        }

        return headers1;
    };

    DocIntSharePoint.prototype.getBearerHeaders = function (bearer, odata) {
        var accept = 'application/json;';

        if (odata) {
            accept = 'application/json;' + odata;
        }
        else {
            //default
            accept += "odata=verbose";
        }
        var headers1 = {
            'Accept': accept,
            'Authorization': bearer
        };

        return headers1;
    };

    DocIntSharePoint.prototype.getSPLibraryAllInfoNoJQuery = function (token) {

        //var listWebEndpoint = "/_api/lists/getbytitle('" + this.list + "')";
        var listWebEndpoint = this.getThisList();
        //"/_api/lists(guid'" + this.listId + "')";

        //https://dtecho365.sharepoint.com/sites/ServiceNow/_api/lists

        var listWebUrl = this.buildSPRestUri(listWebEndpoint);
        var result = {}, headers1, xhrFields;

        if (token != null) {
            var bearer = this.buildAuthToken(token);
            headers1 = this.getBearerHeaders(bearer);
        } else {
            //no token, implies CORS/NTLM Auth
            headers1 = this.getCORSAuthHeaders();
            xhrFields = true;
        }

        var listWebInfo = this.restCallWithoutJQueryV2("GET", listWebUrl, false, xhrFields, headers1);

        if (listWebInfo && !listWebInfo.error && listWebInfo.data) {

            var responseDataObj = JSON.parse(listWebInfo.data);

            if (responseDataObj.d) {
                result.library_id = responseDataObj.d.Id;
            }
            else {
                result.error = "Failed retrieving library list from SharePoint response \n" + listsInfo.data;
            }
        }
        else {
            result.error = listWebInfo.error || listWebInfo.error;
        }

        //try get one drive id
        // note: for sponline, from list retriveal can use /drives single call to get everyhing except item type
        //for sponprem, still use the same call, no need for one drive stuff

        //https://dtecho365.sharepoint.com/sites/ServiceNowDev1/_api/v2.0/drives?select=id,name,description,driveType

        return result;
    };


    DocIntSharePoint.prototype.getSPOnlineLibraryListWithOneDriveIdNoJQuery = function (token) {
        var driveListWebEndpoint = '/_api/v2.0/drives';

        var drivesWebUrl = this.buildSPRestUri(driveListWebEndpoint);
        var result = {}, headers2, xhrFields;

        if (token != null) {
            var bearer = this.buildAuthToken(token);
            //{"error":{"code":"-1, Microsoft.SharePoint.Client.ClientServiceException","message":"The HTTP header ACCEPT is missing or its value is invalid."}}
            //no odata=verbose
            headers2 = this.getBearerHeaders(bearer, "odata.metadata=minimal");
        } else {
            //no token, implies CORS/NTLM Auth
            headers2 = this.getCORSAuthHeaders();
            xhrFields = true;
        }

        var drivesInfo = this.restCallWithoutJQueryV2("GET", drivesWebUrl, false, xhrFields, headers2);

        if (drivesInfo && !drivesInfo.error && drivesInfo.data) {

            var responseDataObj = JSON.parse(drivesInfo.data);

            result.results = responseDataObj.value;
        }
        else {
            result.error = drivesInfo.error;
        }

        return result;
    };

    DocIntSharePoint.prototype.getSPLibraryListNoJQuery = function (token) {
        var listWebEndpoint = '/_api/lists';
        //https://dtecho365.sharepoint.com/sites/ServiceNow/_api/lists

        var listWebUrl = this.buildSPRestUri(listWebEndpoint);
        var result = {}, headers1, xhrFields;

        if (token != null) {
            var bearer = this.buildAuthToken(token);
            headers1 = this.getBearerHeaders(bearer, "odata=verbose");
        } else {
            //no token, implies CORS/NTLM Auth
            headers1 = this.getCORSAuthHeaders();
            xhrFields = true;
        }

        var listsInfo = this.restCallWithoutJQueryV2("GET", listWebUrl, false, xhrFields, headers1);
        if (listsInfo && !listsInfo.error && listsInfo.data) {

            var responseDataObj = JSON.parse(listsInfo.data);

            if (responseDataObj.d && responseDataObj.d.results) {
                result.results = responseDataObj.d.results;
            }
            else {
                result.error = "Failed retrieving library list from SharePoint response \n" + listsInfo.data;
            }
        }
        else {
            result.error = listsInfo.error;
        }

        return result;
    };

    DocIntSharePoint.prototype.validateSPSite = function (token, relative_url) {

        this.site = relative_url;

        var siteCollectionEndpoint = '/_api/site';
        var siteWebEndpoint = '/_api/web';

        var siteCollectionInfoUrl = this.buildSPRestUri(siteCollectionEndpoint);
        var siteWebInfoUrl = this.buildSPRestUri(siteWebEndpoint);

        var result = {}, headers1, xhrFields;

        if (token != null) {
            var bearer = this.buildAuthToken(token);
            headers1 = this.getBearerHeaders(bearer);
        } else {
            //no token, implies CORS/NTLM Auth
            headers1 = this.getCORSAuthHeaders();
            xhrFields = true;
        }

        var siteCollectionInfo = this.restCallWithoutJQueryV2("GET", siteCollectionInfoUrl, false, xhrFields, headers1);
        var siteWebInfo = this.restCallWithoutJQueryV2("GET", siteWebInfoUrl, false, xhrFields, headers1);

        var siteCollectionId, siteWebId, siteId, siteTitle, siteWebUrl;

        if (siteCollectionInfo && !siteCollectionInfo.error && siteWebInfo && !siteWebInfo.error) {
            siteCollectionId = JSON.parse(siteCollectionInfo.data).d.Id;
            var d = JSON.parse(siteWebInfo.data).d;
            siteWebId = d.Id;
            siteTitle = d.Title;
            siteUrl = d.Url;

            var hostname = this.baseUrl.replace('http://', '').replace('https://', '');
            siteId = hostname + "," + siteCollectionId + "," + siteWebId;

            //401, invalid token??? > Grant Sites.Read.All Delegated
            //var siteInfo = this.restCallWithoutJQuery("GET", "https://graph.microsoft.com/v1.0/sites?search=sitename", false, xhrFields, headers1);

            result.siteId = siteId;
            result.siteTitle = siteTitle;
            result.siteUrl = siteUrl;
            result.relative_url = relative_url;
            result.siteDescription = d.Description;
        }
        else {
            result.error = siteCollectionInfo.error || siteWebInfo.error;
        }

        return result;
    };

    DocIntSharePoint.prototype.getAJAXParamWrapper = function (token) {
        var AJAXParamWrapper = {};
        var headers1, xhrFields;

        if (!token.endsWith('-0000')) {
            var bearer = this.buildAuthToken(token);
            headers1 = this.getBearerHeaders(bearer);
            xhrFields = {};
        } else {
            headers1 = this.getCORSAuthHeaders(token);
            xhrFields = {
                withCredentials: true
            };
        }

        AJAXParamWrapper.headers = headers1;
        AJAXParamWrapper.xhrFields = xhrFields;
        return AJAXParamWrapper;
    };

    DocIntSharePoint.prototype.buildSPItemRefRestUri = function (itemRefEndPoint) {

        var url = this.baseUrl + this.site;
        //var url = this.site;

        // remove the trailing /if it exists
        if (url[url.length - 1] == '/') {
            url = url.substring(0, url.length - 1);
        }

        // add the endpoint
        url += '/_api/' + itemRefEndPoint;

        // return the result
        return url;
    };

    DocIntSharePoint.prototype.createTaskIdMeta = function (name, value) {
        var meta = '{\'__metadata\':{\'type\':\'SP.ListItem\'}, \'' + name + '\':\'' + value + '\'}';

        return meta;
    };

    DocIntSharePoint.prototype.checkStatus = function (response) {
        if (response.status >= 200 && response.status < 300) {
            return Promise.resolve(response);
        } else {
            return Promise.reject(new Error(response.statusText));
        }
    };

    DocIntSharePoint.prototype.parseJson = function (response) {
        return response.json();
    };

    DocIntSharePoint.prototype.getDigestToken = function (root) {
        var digestToken;

        var endpoint = '/_api/contextinfo';
        var restUrl = this.buildSPRestUri(endpoint);

        if (root) {
            restUrl = this.baseUrl + endpoint;
        }

        $.ajax({
            url: restUrl,
            type: "POST",
            crossDomain: true,
            async: false,
            headers: this.getCORSAuthHeaders(),
            xhrFields: {
                withCredentials: true
            },
            success: function (data, status) {
                digestToken = data.d.GetContextWebInformation.FormDigestValue;
            },
            error: function (xhr, e, error) {

                if (error.name == "NetworkError" && !xhr.responseText) {
                    alert("Network Error to SharePoint, Please check if SharePoint Server is up.");
                }
                else {
                    alert('Error getting digest token [' + error + "], Response: " + xhr.responseText);
                }
            }
        });

        return digestToken;
    };

    DocIntSharePoint.prototype.getDigestTokenNoJQuery = function (root) {
        var digestToken;
        var endpoint = '/_api/contextinfo';
        var restUrl = this.buildSPRestUri(endpoint);

        if (root) {
            restUrl = this.baseUrl + endpoint;
        }

        var result, headers1, xhrFields;

        headers1 = this.getCORSAuthHeaders();
        xhrFields = true;

        var data = this.restCallWithoutJQuery("POST", restUrl, false, xhrFields, headers1);

        if (data && JSON.parse(data).d.GetContextWebInformation.FormDigestValue) {
            digestToken = JSON.parse(data).d.GetContextWebInformation.FormDigestValue;
            alert("Connection Test Succeeds!");
        } else {
            alert("Connection Test fails");
        }

        return digestToken;
    };

    DocIntSharePoint.prototype.getSPContentTypeFieldsNoJQuery = function (token, contentTypeId) {
        //ContentTypes('0x0101')/Fields
        var endpoint = this.getThisList() + "/ContentTypes('" + contentTypeId + "')/Fields";
        var restUrl = this.buildSPRestUri(endpoint);
        var result, headers1, xhrFields;

        if (token != null) {
            var bearer = this.buildAuthToken(token);
            headers1 = this.getBearerHeaders(bearer);
        } else {
            headers1 = this.getCORSAuthHeaders();
            xhrFields = true;
        }

        return this.restCallWithoutJQuery("GET", restUrl, false, xhrFields, headers1);
    };

    DocIntSharePoint.prototype.getSPContentTypeNoJQuery = function (token) {
        var endpoint = this.getThisList() + "/ContentTypes";
        var restUrl = this.buildSPRestUri(endpoint);
        var result, headers1, xhrFields;

        if (token != null) {
            var bearer = this.buildAuthToken(token);
            headers1 = this.getBearerHeaders(bearer);
        } else {
            headers1 = this.getCORSAuthHeaders();
            xhrFields = true;
        }

        return this.restCallWithoutJQuery("GET", restUrl, false, xhrFields, headers1);
    };

    DocIntSharePoint.prototype.getSPViewFieldsNoJQuery = function (token) {
        var endpoint = this.getThisList() + "/fields";
        var restUrl = this.buildSPRestUri(endpoint);
        var result, headers1, xhrFields;

        if (token != null) {

            var bearer = this.buildAuthToken(token);
            headers1 = this.getBearerHeaders(bearer);

        } else {
            headers1 = this.getCORSAuthHeaders();
            xhrFields = true;
        }

        return this.restCallWithoutJQuery("GET", restUrl, false, xhrFields, headers1);
    };

    DocIntSharePoint.prototype.testLibraryNoJQuery = function (token) {

        var endpoint = this.getThisList() + "/";

        var restUrl = this.buildSPRestUri(endpoint);
        var headers1, xhrFields;

        if (token != null) {
            var bearer = this.buildAuthToken(token);
            headers1 = this.getBearerHeaders(bearer);

            xhrFields = {};

        } else {
            headers1 = this.getCORSAuthHeaders();

            xhrFields = {
                withCredentials: true
            };
        }

        var result = this.restCallWithoutJQueryV2("GET", restUrl, false, xhrFields, headers1);

        // if (data && JSON.parse(data).d) {
        //     result = JSON.parse(data).d;
        // }

        return result;
    };

    DocIntSharePoint.prototype.getSPQuery = function (endpoint, token) {

        var restUrl = this.buildSPRestUri(endpoint);
        var result = {},
            headers1, xhrFields;

        if (token != null && !token.endsWith('-0000')) {
            var bearer = this.buildAuthToken(token);
            headers1 = this.getBearerHeaders(bearer);
            xhrFields = {};
        } else {
            headers1 = this.getCORSAuthHeaders();
            xhrFields = {
                withCredentials: true
            };
        }

        $.ajax({
            url: restUrl,
            type: "GET",
            dataType: 'JSON',
            async: false,
            headers: headers1,
            xhrFields: xhrFields,
            success: function (responseData) {
                result = responseData;
            },
            error: function (xhr, e, error) {
                if (xhr.responseJSON.error.code == "-2147024860, Microsoft.SharePoint.SPQueryThrottledException") {
                    alert("\nSharePoint Library reached maxiumn number of document that query can reach, please adjust and try it again\n\n" + 'Internal SharePoint Query Error ' + JSON.stringify(xhr.responseJSON.error, null, 2));
                } else if (xhr.responseJSON.error) {
                    alert('SharePoint Query Error ' + xhr.responseJSON.error.code + ", message " + xhr.responseDataObj.error.message);
                } else {
                    alert('SharePoint query Error ' + xhr.responseJSON);
                }
            }
        });

        return result;
    };

    DocIntSharePoint.prototype.getSPListItems = function (token) {
        var endpoint =  this.getThisList()  + "/items";
        var result = this.getSPQuery(endpoint, token);

        return result;
    };

    DocIntSharePoint.prototype.getSPListItemsNoJQuery = function (token) {
        var endpoint = this.getThisList() + "/items";
        var result = this.getSPQueryNoJQuery(endpoint, token);

        return result;
    };
    // DocIntSharePoint.prototype.testLibraryNoJQuery = function (token) {
    //     var result = this.runGetListNoJQuery(token);
    //     var libaryResult = {};
    //     var result = {};

    //     var message;

    //     if (resultD) {
    //         message = 'Test Library Succeeds!';

    //         libaryResult.sp_library_id = resultD.Id;
    //         libaryResult.sp_item_type = resultD.ListItemEntityTypeFullName;
    //     }
    //     else {
    //         message = 'Test Library Failed!';
    //     }

    //     result.message = message;
    //     result.libaryResult = libaryResult;
    //     result.error = 

    //     return result;
    // };

    DocIntSharePoint.prototype.getSPDocumentsByContentType = function (token, filterStr, finalOrderBy) {
        var endpoint;
        var expandStatement;

        if (this.sharePointVersion == undefined || this.sharePointVersion == '' || this.sharePointVersion == 'O365') {
            expandStatement = "$select=*,ServerRedirectedEmbedUri,ContentType&$expand=FieldValuesAsText,ContentType";

        }
        else {
            //else if (this.sharePointVersion == 'SP2013-NTLM' || this.sharePointVersion == 'SP2016-NTLM' || this.sharePointVersion == 'SP2019-NTLM') {
            expandStatement = '$select=*,ContentType,FieldValuesAsText,Author/Id&$expand=ContentType,FieldValuesAsText/Id,Author/Title,Author/Id,Author';
        }

        if (filterStr === undefined || filterStr == '') {
            //endpoint = '/_api/web/Lists/GetByTitle(\'' + this.list + '\')/items?' + expandStatement;
            endpoint = this.getThisList() + "/items?filter=(ID eq -1)&" + expandStatement;

        } else {
            endpoint = this.getThisList() + "/items?$filter=(" + filterStr + ")&" + expandStatement;
        }

        if (finalOrderBy != "") {
            endpoint = endpoint + "&" + finalOrderBy;
        }

        var result = this.getSPQuery(endpoint, token);
        if (result && !$.isEmptyObject(result)) {
            this.processMapping(result, token, this.sharePointVersion);
        }
        return result;
    };

    //Used by DocIntegrator by full filter
    DocIntSharePoint.prototype.getSPDocumentsByFilter = function (token, filterStr, finalOrderBy) {
        var endpoint;
        var expandStatement;

        if (this.sharePointVersion == undefined || this.sharePointVersion == '' || this.sharePointVersion == 'O365') {
            expandStatement = '$expand=FieldValuesAsText&$select=ServerRedirectedEmbedUri,*';
        }
        //else if (this.sharePointVersion == 'SP2013-NTLM' || this.sharePointVersion == 'SP2016-NTLM' || this.sharePointVersion == 'SP2019-NTLM') {
        else {
            expandStatement = '$select=*,FieldValuesAsText,Author/Id&$expand=FieldValuesAsText/Id,Author/Title,Author/Id,Author';
        }

        if (!filterStr) {
            filterStr = "ID eq -1";
        }

        endpoint = this.getThisList() + "/items?$filter=(" + filterStr + ")&" + expandStatement;

        if (finalOrderBy) {
            endpoint = endpoint + "&" + finalOrderBy;
        }

        var result = this.getSPQuery(endpoint, token);

        this.processMapping(result, token, this.sharePointVersion);

        return result;
    };

    DocIntSharePoint.prototype.processMapping = function (result, token, version) {
        var authorEndpoint = this.getThisList() + "/items?$select=Id,Author/Id,Author/Title,Editor/Id,Editor/Title&$expand=Author,Editor";
        var authorResult = this.getSPQuery(authorEndpoint, token);

        for (var r in result.d.results) {
            for (var a in authorResult.d.results) {
                if (result.d.results[r].FieldValuesAsText.Author == authorResult.d.results[a].Author.Id || result.d.results[r].AuthorId == authorResult.d.results[a].Author.Id && result.d.results[r].FieldValuesAsText.ID == authorResult.d.results[a].ID) {
                    result.d.results[r].FieldValuesAsText.Author = authorResult.d.results[a].Author.Title;
                    result.d.results[r].FieldValuesAsText.Editor = authorResult.d.results[a].Editor.Title;

                    //set up for content type view fields which don't have Author/Editor
                    result.d.results[r].FieldValuesAsText.Created_x0020_By = authorResult.d.results[a].Author.Title;
                    result.d.results[r].FieldValuesAsText.Modified_x0020_By = authorResult.d.results[a].Editor.Title;
                }
            }
        }
    };
    //Retired by DocIntegrator
    DocIntSharePoint.prototype.getSPDocumentsByItemId = function (token, item, itemId) {
        var endpoint =  this.getThisList() + "/items?$filter=(" + item + " eq '" + itemId + "')&$expand=FieldValuesAsText&$select=ServerRedirectedEmbedUri,*";
        var result = this.getSPQuery(endpoint, token);

        return result;
    };

    DocIntSharePoint.prototype.search = function (q, token, searchContext) {
        //token will be null for on-prem

        var urlProps;
        var scope = searchContext.scope;
        var searchQuery;

        //limit by site collection
        var sitePath;

        //limit by library
        var list;

        var queryTextOpen = "{'request': { 'Querytext':'" + q;
        var queryTextClose = "',";
        var selectProps = searchContext.select_properties;
        var postAPI = "/_api/search/postquery";

        //urlProps = searchContext.conn.baseURL + postAPI;

        //default to site-aware URL, 
        urlProps = searchContext.conn.baseURL + searchContext.conn.relative_url + postAPI;
        var rootLevel = false;

        sitePath = " path:" + searchContext.conn.baseURL + searchContext.conn.relative_url;

        if (scope == "connection") {
            searchQuery = queryTextOpen + queryTextClose + selectProps;
            //remove site-level
            urlProps = searchContext.conn.baseURL + postAPI;
            rootLevel = true;
        } else if (scope == "site") {
            searchQuery = queryTextOpen + sitePath + queryTextClose + selectProps;

        } else if (scope == "library") {
            searchQuery = queryTextOpen + sitePath + "/" + searchContext.conn.list_title + queryTextClose + selectProps;
        } else if (scope == "content_type") {
            var contentType = searchContext.conn.content_type;
            var contentTypeFilter = '';

            if (contentType != null && contentType != 'ALL') {
                contentTypeFilter = ' ContentType:' + contentType;
            }

            var content_type_search_scope = searchContext.conn.content_type_scope;

            if (content_type_search_scope == 'Site') {
                searchQuery = queryTextOpen + contentTypeFilter + sitePath + queryTextClose + selectProps;
            } else if (content_type_search_scope == 'Library') {
                searchQuery = queryTextOpen + contentTypeFilter + sitePath + "/" + searchContext.conn.list_title + queryTextClose + selectProps;
            } else {
                //connection level                
                searchQuery = queryTextOpen + contentTypeFilter + queryTextClose + selectProps;

                urlProps = searchContext.conn.baseURL + postAPI;
                rootLevel = true;
            }
        } else {
            var configErrMsg = "Unsupported Search Scope [" + scope + "]";

            alert(configErrMsg);
        }

        var resultItems = '';
        var headers1, headers2, xhrFields;

        if (!token) {
            //site level or connection level token is different for search for on-prem case
            token = this.getDigestToken(rootLevel);
        }

        if (token && !token.endsWith('-0000')) {

            var bearer = this.buildAuthToken(token);

            headers2 = {
                'Accept': 'application/json; odata=verbose',
                'Authorization': bearer,
                'Content-Type': 'application/json; odata=verbose'
            };

            xhrFields = {};
        } else {
            headers2 = this.getCORSAuthHeaders(token);
            xhrFields = {
                withCredentials: true
            };
        }

        $.ajax({
            url: urlProps,
            type: 'POST',
            //crossDomain: true,
            async: false,
            headers: headers2,
            data: searchQuery,
            processData: false,
            dataType: 'json',
            xhrFields: xhrFields,
            success: function (data, status) {
                resultItems = data.d.postquery.PrimaryQueryResult.RelevantResults;
            },
            error: function (xhr, e, error) {
                alert("Search error: " + JSON.stringify(error) + "\r\ne: " + JSON.stringify(e) + "xhr.responseText: " + JSON.stringify(xhr.responseText));
            }
        });

        return resultItems;
    };

    DocIntSharePoint.prototype.postSPDocumentMeta = function (itemRef, token, meta) {
        var itemId = -1;
        var urlProps = itemRef + '/listitemallfields';
        var urlListItem = '';

        var headers1, headers2, xhrFields;

        if (!token.endsWith('-0000')) {
            var bearer = this.buildAuthToken(token);
            headers1 = this.getBearerHeaders(bearer);

            headers2 = {
                'Accept': 'application/json; odata=verbose',
                'Authorization': bearer,
                'Content-Type': 'application/json; odata=verbose',
                'IF-MATCH': '*',
                'X-HTTP-Method': 'MERGE'
            };

            xhrFields = {};
        } else {

            if (this.sharePointVersion.toUpperCase() == 'SP2013-NTLM') {
                //before urlProps = "Web/api/abcd";
                urlProps = this.buildSPItemRefRestUri(urlProps);
                //after urlProps = "https://host/sites/sitename/Web/api/abcd"
            }

            headers1 = {
                "Content-Type": "'application/json;odata=verbose'",
                "Accept": "application/json;odata=verbose",
                "crossDomain": "true",
                "credentials": "include" //,
            };

            headers2 = {
                "Content-Type": "application/json;odata=verbose",
                "Accept": "application/json;odata=verbose",
                "crossDomain": "true",
                "credentials": "include",
                'IF-MATCH': '*',
                'X-HTTP-Method': 'MERGE',
                "X-RequestDigest": token
            };

            xhrFields = {
                withCredentials: true
            };
        }

        // get the document list item id
        $.ajax({
            url: urlProps,
            type: "GET",
            crossDomain: true,
            async: false,
            headers: headers1,
            xhrFields: xhrFields,
            success: function (data, status) {
                // get the item id
                itemId = data.d.ID;
                urlListItem = data.d.__metadata.uri;
            },
            error: function (xhr, e, error) {
                alert('error' + error + e + xhr);
            }
        });

        this.postMetadataByItemUri(token, urlListItem, meta);
    };

    DocIntSharePoint.prototype.postMetadataByItemUri = function (token, itemUri, meta) {
        // update the list item 
        var headers2, xhrFields;

        if (!token.endsWith('-0000')) {
            var bearer = this.buildAuthToken(token);
            headers2 = {
                'Accept': 'application/json; odata=verbose',
                'Authorization': bearer,
                'Content-Type': 'application/json; odata=verbose',
                'IF-MATCH': '*',
                'X-HTTP-Method': 'MERGE'
            };

            xhrFields = {};
        } else {
            headers2 = {
                "Content-Type": "application/json;odata=verbose",
                "Accept": "application/json;odata=verbose",
                "crossDomain": "true",
                "credentials": "include",
                'IF-MATCH': '*',
                'X-HTTP-Method': 'MERGE',
                "X-RequestDigest": token
            };

            xhrFields = {
                withCredentials: true
            };
        }

        var updateSuccess = false;

        $.ajax({
            url: itemUri,
            type: 'POST',
            crossDomain: true,
            async: false,
            data: meta,
            length: meta.length,
            headers: headers2,
            xhrFields: xhrFields,
            success: function (data, status) {
                updateSuccess = true;
            },
            error: function (xhr, e, error) {
                alert('Metadata Updating Error \n' + xhr.responseJSON.error.message.value + "\r\nMeta:  " + meta);
            }
        });

        return updateSuccess;
    };

    DocIntSharePoint.prototype.checkFilenameIsExistBySPSiteURL = function (token, SPSiteURL, fileByServerRelativeUrl, AJAXParamWrapper) {   //Short GET Calls to check if a file name is being used. 
        //fileByServerRelativeUrl => '/sites/SerivceNowDev2/targetLibname/filename
        var restUrl = SPSiteURL + '/_api/web/getfilebyserverrelativeurl(\'' + fileByServerRelativeUrl + '\')';

        var isExist = false;
        $.ajax({
            url: restUrl,
            type: "GET",
            crossDomain: true,
            async: false,
            headers: AJAXParamWrapper.headers,
            xhrFields: AJAXParamWrapper.xhrFields,
            success: function (data, status) {
                //result = data.d.__metadata.id;
                isExist = true;
            },
            error: function (xhr, e, error) {
                if (error.status == 404) {
                    //log("file doesnt exist");
                }
            }
        });

        return isExist;
    };

    DocIntSharePoint.prototype.copySP2SP = function (token, item, contextObj, searchContext) {
        var AJAXParamWrapper = this.getAJAXParamWrapper(token);
        var sourceConn = searchContext.conn;
        var targetConn = contextObj.conn;

        //file1        
        var filename = item.Filename;
        //https://hostname/sites/servicenow/Dev1Lib1/test-folder1/canvas.pdf"

        var sitesIndexPos = item.Path.indexOf("/sites/");
        var filenameIndexPos = item.Path.indexOf(filename);

        // ServerRelativeUrl without filename        
        var relativePathTarget = targetConn.relative_url + "/" + targetConn.list_title + "/";

        //"/sites/ServiceNowDev1/Dev1Lib1/" or //"/sites/ServiceNowDev1/Dev1Lib1/DS-folder"
        var relativePathSource = item.Path.substring(sitesIndexPos, filenameIndexPos);

        var isFileExist = true;
        var temp_fileByServerRelativeUrl = relativePathTarget + filename;

        if (relativePathSource != relativePathTarget) {
            //could be the same lib due to the folder or different lib
            var targetSitURL = targetConn.baseURL + targetConn.relative_url;

            isFileExist = this.checkFilenameIsExistBySPSiteURL(token, targetSitURL, temp_fileByServerRelativeUrl, AJAXParamWrapper);
        }
        else {
            //same lib, same folder (root level)
        }

        var temp_new_file_name;

        if (isFileExist) {
            var i = 1;
            var name_only = filename.split('.').slice(0, -1).join('.');
            var ext = filename.split('.').pop();

            if (ext == filename) {
                ext = '';
                name_only = filename;
            }

            //continue to check until found one that doesn't exist.
            while (isFileExist) {
                temp_new_file_name = name_only + " " + i;
                i++;

                if (i > 2000) break;

                if (ext && ext.length > 0) {
                    temp_new_file_name += "." + ext;
                }

                temp_fileByServerRelativeUrl = relativePathTarget + temp_new_file_name;
                isFileExist = this.checkFilenameIsExistBySPSiteURL(token, item.SPSiteURL, temp_fileByServerRelativeUrl, AJAXParamWrapper);
            }

            if (temp_new_file_name) {
                filename = temp_new_file_name;
            }
        }

        var urlCopyTo = item.SPSiteURL + "/_api/web/getfilebyserverrelativeurl('" + item.Path.substring(sitesIndexPos) + "')/copyTo(strNewUrl='" + temp_fileByServerRelativeUrl + "',bOverWrite=false)";

        var fileCopySuccess = false;
        //Make a copy of the document
        $.ajax({
            url: urlCopyTo,
            type: 'POST',
            crossDomain: true,
            async: false,
            //data: meta,
            //length: meta.length,
            headers: AJAXParamWrapper.headers,
            xhrFields: AJAXParamWrapper.xhrFields,
            success: function (data, status) {
                //alert('Docment copied Successfully to ' + temp_fileByServerRelativeUrl);
                //            
                fileCopySuccess = true;
            },
            error: function (xhr, e, error) {
                alert('File fails to copy - Error \n' + xhr.responseJSON.error.message.value);
            }
        });

        var itemUri;
        var itemRef = item.SPSiteURL + "/_api/web/getfilebyserverrelativeurl('" + temp_fileByServerRelativeUrl + "')";

        if (fileCopySuccess) {
            var urlProps = itemRef + '/listitemallfields';

            // get the document list item id
            $.ajax({
                url: urlProps,
                type: "GET",
                crossDomain: true,
                async: false,
                headers: AJAXParamWrapper.headers,
                xhrFields: AJAXParamWrapper.xhrFields,
                success: function (data, status) {
                    // get the item id
                    itemUri = data.d.__metadata.uri;
                },
                error: function (xhr, e, error) {
                    alert('error' + error + e + xhr);
                }
            });

        }
        else {
            alert("File Transfer fails");
        }

        if (itemUri) {
            itemUri = itemUri.replace(/[{}]/g, "");
        }
        return itemUri;
    };

    // https://social.msdn.microsoft.com/Forums/office/en-US/357f5c7d-77b1-4d04-9956-3498ce33a5e0/upload-big-file-to-sharepoint-with-rest-api
    // TODO: Need to look at adding a parameter for folder specification.
    DocIntSharePoint.prototype.postSPFile = function (token, filename, buffer, meta, folder) {
        //upload content
        //var itemRef = this.postSPFileBinary(token, filename, buffer, "/testfolder1/testfolder2");
        var itemRef = this.postSPFileBinary(token, filename, buffer, folder);

        //update metadata
        if (itemRef) {
            this.postSPDocumentMeta(itemRef, token, meta);
        }

        return itemRef;
    };

    DocIntSharePoint.prototype.postSPFileBinary = function (token, filename, buffer, folder) {
        //fetch common headers
        var AJAXParamWrapper = this.getAJAXParamWrapper(token);

        //handles binary file upload and duplicate file names
        var isFileExist = this.checkFilenameIsExist(token, filename, AJAXParamWrapper, folder);

        if (isFileExist) {
            var i = 1;
            var name_only = filename.split('.').slice(0, -1).join('.');
            var ext = filename.split('.').pop();

            if (ext == filename) {
                ext = '';
                name_only = filename;
            }
            var temp_new_file_name;

            //continue to check until found one that doesn't exist.
            while (isFileExist) {
                temp_new_file_name = name_only + " " + i;
                i++;

                if (i > 1000) break;

                if (ext && ext.length > 0) {
                    temp_new_file_name += "." + ext;
                }

                isFileExist = this.checkFilenameIsExist(token, temp_new_file_name, AJAXParamWrapper, folder);
            }

            if (temp_new_file_name) {
                filename = temp_new_file_name;
            }
        }

        //var endpoint = '/_api/web/lists/(\'' + this.list + '\')/rootfolder/';
        var endpoint = "/_api/web/GetFolderByServerRelativeUrl(\'" + this.site + "/" + this.listPath;
        // + folder + "\')";

        if (folder) {
            endpoint += folder;
        }
        else {
        }

        endpoint += "\')/files/add(url=\'" + filename + "\', overwrite = false)";

        var restUrl = this.buildSPRestUri(endpoint);

        var itemRef;
        $.ajax({
            url: restUrl,
            type: "POST",
            crossDomain: true,
            data: buffer,
            async: false,
            processData: false,
            length: buffer.byteLength,
            headers: AJAXParamWrapper.headers,
            xhrFields: AJAXParamWrapper.xhrFields,
            success: function (data, status) {
                //result = data.d.__metadata.id;
                itemRef = data.d.__metadata.id;

                //for 2013, itemRef looks like 'Web/a/b/c
                //for 2016/2019, itemRef looks like 'http://host/site/sitename/_api/Web/a/b/c                
            },
            error: function (xhr, e, error) {
                alert('Upload SharePoint Error ' + error + e + xhr.responseText);
            }
        });

        return itemRef;
    };

    DocIntSharePoint.prototype.checkFilenameIsExist = function (token, filename, AJAXParamWrapper, folder) {
        var fileByServerRelativeUrl = this.site + "/" + this.listPath;

        if (folder) {
            fileByServerRelativeUrl += folder;
        }

        fileByServerRelativeUrl += "/" + filename;

        var endpoint = '/_api/web/getfilebyserverrelativeurl(\'' + fileByServerRelativeUrl + '\')';

        var restUrl = this.buildSPRestUri(endpoint);
        var isExist = false;

        $.ajax({
            url: restUrl,
            type: "GET",
            crossDomain: true,
            async: false,
            headers: AJAXParamWrapper.headers,
            xhrFields: AJAXParamWrapper.xhrFields,
            success: function (data, status) {
                //result = data.d.__metadata.id;
                isExist = true;
            },
            error: function (xhr, e, error) {
                if (error.status == 404) {
                    //log("file doesnt exist");
                }
            }
        });

        return isExist;
    };

    DocIntSharePoint.prototype.postSPFileNoJQuery = function (token, filename, buffer, meta) {
        var endpoint = this.getThisList() + "/rootfolder/files/add(url='" + filename + "', overwrite = true)";
        var restUrl = this.buildSPRestUri(endpoint);
        var result = null;
        var localRef = this;
        var headers1, xhrFields;

        if (!token.endsWith('-0000')) {
            var bearer = this.buildAuthToken(token);
            headers1 = this.getBearerHeaders(bearer);
            xhrFields = {};
        } else {
            var headers_temp = this.getCORSAuthHeaders();
            headers_temp["X-RequestDigest"] = token;
            headers1 = headers_temp;
            xhrFields = {
                withCredentials: true
            };
        }

        var xhr = new XMLHttpRequest();

        xhr.open("POST", restUrl, false);
        // `false` makes the request synchronous
        //set headers
        for (var key in headers1) {
            xhr.setRequestHeader(key, headers1[key]);
        }

        //set cors
        xhr.withCredentials = xhrFields;

        xhr.send(buffer);

        if (xhr.status === 200) {
            jslog(xhr.responseText);

            return xhr.responseText;
        } else {
            return {};
        }
    };

    DocIntSharePoint.prototype.getLibContentTypes = function (token) {
        var endpoint = this.getThisList() + "/contenttypes";
        var restUrl = this.buildSPRestUri(endpoint);
        var result, headers1, xhrFields;

        if (token != null) {

            var bearer = this.buildAuthToken(token);
            headers1 = this.getBearerHeaders(bearer);

            xhrFields = {};

        } else {
            headers1 = this.getCORSAuthHeaders();

            xhrFields = {
                withCredentials: true
            };
        }
        $.ajax({
            url: restUrl,
            type: "GET",
            dataType: 'JSON',
            async: false,
            headers: headers1,
            xhrFields: xhrFields,
            success: function (responseData) {
                result = responseData.d.results;
            },
            error: function (xhr, e, error) {
                alert('SharePoint Query Error ' + error + e + xhr.responseText);
            }
        });

        return result;
    };

    DocIntSharePoint.prototype.restCallWithoutJQueryV2 = function (httpMethod, restUrl, asynchronous, withCredentials, headers1) {
        var xhr = new XMLHttpRequest();

        var result = {};

        xhr.onreadystatechange = function () {
            if (xhr.readyState === 4) { //if complete
                if (xhr.status === 200) { //check if "OK" (200)
                    //success
                } else {
                }
            } else {
                //alert('other state');
            }
        };

        xhr.open(httpMethod, restUrl, asynchronous); // `false` makes the request synchronous
        //set headers
        //xhr.setRequestHeader("Cache-Control", "no-cache");

        for (var key in headers1) {
            xhr.setRequestHeader(key, headers1[key]);
        }

        //set cors
        xhr.withCredentials = withCredentials;

        try {
            xhr.send(null);
            if (xhr.status === 200) {
                jslog(xhr.responseText);
                result.data = xhr.responseText;
            } else {
                result.error = xhr.responseText + ", " + xhr.responseURL;
            }
        } catch (e) {
            result.error = e;
        }

        return result;
    };

    DocIntSharePoint.prototype.restCallWithoutJQuery = function (httpMethod, restUrl, asynchronous, withCredentials, headers1) {
        var xhr = new XMLHttpRequest();

        xhr.onreadystatechange = function () {
            if (xhr.readyState === 4) { //if complete
                if (xhr.status === 200) { //check if "OK" (200)
                    //success
                } else {

                    alert('REST Call Error \n' + xhr.responseText + ", " + xhr.responseURL); //otherwise, some other code was returned
                    return null;
                }
            } else {
                //alert('other state');
            }
        };

        xhr.open(httpMethod, restUrl, asynchronous); // `false` makes the request synchronous
        //set headers
        for (var key in headers1) {
            xhr.setRequestHeader(key, headers1[key]);
        }

        //set cors
        xhr.withCredentials = withCredentials;

        try {
            xhr.send(null);
            if (xhr.status === 200) {
                jslog(xhr.responseText);

                return xhr.responseText;
            } else {
                return {};
            }
        } catch (e) {
            alert(" error " + e);
            return null;
        }
    };

    DocIntSharePoint.prototype.processSearchResults = function (relevantResults, serviceNowLinkProperty) {
        //items will be RelevantResults
        var uniqueSIDSet = new Set();
        var listIdArray = [];
        var results = [];
        var processingResults = {};

        var tableRowResults = relevantResults.Table.Rows.results;
        for (var spResultItem in tableRowResults) {
            var searchResultItem = {};
            var DocItem = tableRowResults[spResultItem];

            for (var cellResultsItem in DocItem.Cells.results) {
                var fieldName = DocItem.Cells.results[cellResultsItem].Key;
                var fieldValue = DocItem.Cells.results[cellResultsItem].Value;
                var fieldType = DocItem.Cells.results[cellResultsItem].ValueType;
                var metadata = DocItem.Cells.results[cellResultsItem].__metadata;

                searchResultItem[fieldName] = fieldValue;
                searchResultItem.type = fieldType;
                searchResultItem.metadata = metadata;
            }

            //Create a ServiceNow App ID based on full path
            searchResultItem.SID = searchResultItem.Path;

            uniqueSIDSet.add(searchResultItem.SID);

            //remove {} for SP2013 search
            searchResultItem.ListID = searchResultItem.ListID.replace(/[{}]/g, "");

            if (searchResultItem.ListID && !listIdArray.includes(searchResultItem.ListID)) {
                listIdArray.push(searchResultItem.ListID);
            }

            var itemUri = searchResultItem.SPSiteURL + "/_api/Web/Lists(guid'" + searchResultItem.ListID + "')/Items(" + searchResultItem.ListItemID + ")";

            searchResultItem.itemUri = itemUri;

            // if (searchResultItem.ServerRedirectedURL) {
            //     searchResultItem.url = searchResultItem.ServerRedirectedURL;
            // } else {
            //     searchResultItem.url = searchResultItem.Path;
            // }

            //for PDF online, ServerRedirectedEmbedURL is not empty but ServerRedirectedURL is
            //so ServerRedirectedEmbedURL is more reliable
            if (searchResultItem.ServerRedirectedEmbedURL) {
                searchResultItem.url = searchResultItem.ServerRedirectedEmbedURL;
            } else {
                searchResultItem.url = searchResultItem.Path;
            }

            // set servicenowlink 
            //replace(/'/g,'');
            if (serviceNowLinkProperty) {
                var propertyNameWithoutSingleQuote = serviceNowLinkProperty.replace(/'/g, '');
                if (searchResultItem[propertyNameWithoutSingleQuote]) {
                    searchResultItem.servicenowLink = searchResultItem[propertyNameWithoutSingleQuote];
                }
            }

            results.push(searchResultItem);
        }

        processingResults.results = results;
        processingResults.searchUniqueIDSet = uniqueSIDSet;
        processingResults.hasListIDs = listIdArray;

        return processingResults;
    };

    DocIntSharePoint.prototype.lookupContentTypeId = function (content_type, content_type_blob) {
        var contentTypeBlobObj = JSON.parse(content_type_blob);

        for (var item in contentTypeBlobObj) {
            if (contentTypeBlobObj[item].Name === content_type) {
                return contentTypeBlobObj[item].ContentTypeId;
            }
        }
    };

    DocIntSharePoint.prototype.keepQueryContentSearchItems = function (searchResultsItems, queryResultsItemsMap, selectPropertiesMap, searchResultSIDSet) {
        var results = [];

        for (var n in searchResultsItems) {
            var searchResultItem = searchResultsItems[n];

            var matchingQueryItem = queryResultsItemsMap[searchResultItem.SID];

            if (matchingQueryItem) {
                //preferred to use query item if matched
                results.push(matchingQueryItem);
            } else {
                var contentTypeField;

                Object.keys(searchResultItem).forEach(function (searchResultIndexKey) {
                    //looping through all the keys
                    //e.g. searchResultIndexKey is 'Filename'
                    contentTypeField = selectPropertiesMap[searchResultIndexKey];

                    //for rank, contentTypeField will be null
                    if (contentTypeField && contentTypeField != searchResultIndexKey) {
                        //meaning this contentTypeField is mapped to an index name

                        searchResultItem[contentTypeField] = searchResultItem[searchResultIndexKey];
                        //set the value to a new key
                    }
                });

                var Path = searchResultItem["Path"];
                var SPSiteURL = searchResultItem["SPSiteURL"];

                //position of the slash immediately after the Sit URL which is the length of Site URL
                var indexOfFirstSlash = SPSiteURL.length + 1;

                var pathWithoutSite = Path.substring(indexOfFirstSlash);

                var indexOfSecondSlash = pathWithoutSite.indexOf('/');

                var list_title = pathWithoutSite.substring(0, indexOfSecondSlash);

                searchResultItem["edit_url"] = SPSiteURL + "/" + list_title + "/Forms/EditForm.aspx?ID=" + searchResultItem.ListItemID + "&IsDlg=1";

                //see line 1334 about ServerRedirectedEmbedURL
                if (searchResultItem.ServerRedirectedEmbedURL) {
                    searchResultItem["url"] = searchResultItem.ServerRedirectedEmbedURL;
                } else {
                    searchResultItem["url"] = Path;
                }

                results.push(searchResultItem);
            }
        }

        return results;
    };

    DocIntSharePoint.prototype.mergeAndProcessSPCombinedItems = function (searchResultsItems, queryResultsItems, selectPropertiesMap, querySIDSet) {
        //var mergedArray = searchResultsItems.concat(queryResultsItems);

        for (var n in searchResultsItems) {
            var searchResultItem = searchResultsItems[n];

            if (querySIDSet.has(searchResultItem.SID)) {
                //alert("duplicate removed " + searchResultItem.SID);
            } else {
                //
                var contentTypeField;

                Object.keys(searchResultItem).forEach(function (searchResultIndexKey) {
                    //looping through all the keys
                    //e.g. searchResultIndexKey is 'Filename', contentTypeField is 'FileLeafRef'
                    contentTypeField = selectPropertiesMap[searchResultIndexKey];

                    //for rank, contentTypeField will be null
                    if (contentTypeField && contentTypeField != searchResultIndexKey) {
                        //meaning this contentTypeField is mapped to an index name

                        searchResultItem[contentTypeField] = searchResultItem[searchResultIndexKey];
                        //set the value to a new key
                    }
                });

                var Path = searchResultItem["Path"];
                var SPSiteURL = searchResultItem["SPSiteURL"];

                //position of the slash immediately after the Sit URL which is the length of Site URL
                var indexOfFirstSlash = SPSiteURL.length + 1;

                var pathWithoutSite = Path.substring(indexOfFirstSlash);

                var indexOfSecondSlash = pathWithoutSite.indexOf('/');

                var list_title = pathWithoutSite.substring(0, indexOfSecondSlash);

                searchResultItem["edit_url"] = SPSiteURL + "/" + list_title + "/Forms/EditForm.aspx?ID=" + searchResultItem.ListItemID + "&IsDlg=1";

                if (searchResultItem.ServerRedirectedEmbedURL) {
                    searchResultItem["url"] = searchResultItem.ServerRedirectedEmbedURL;
                } else {
                    searchResultItem["url"] = Path;
                }

                queryResultsItems.push(searchResultItem);
            }
        }

        return queryResultsItems;
    };

    return DocIntSharePoint;
}());
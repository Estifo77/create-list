var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : new P(function (resolve) { resolve(result.value); }).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g;
    return g = { next: verb(0), "throw": verb(1), "return": verb(2) }, typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (_) try {
            if (f = 1, y && (t = y[op[0] & 2 ? "return" : op[0] ? "throw" : "next"]) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [0, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
import { SPHttpClient, } from "@microsoft/sp-http";
// import reject = Promise.reject;
var getHeader = {
    headers: {
        accept: "application/json;",
    },
};
var postHeader = {
    headers: {
        "content-type": "application/json;odata.metadata=full",
        accept: "application/json;odata.metadata=full",
    },
};
var deleteHeader = {
    headers: {
        "content-type": "application/json;odata.metadata=full",
        "IF-MATCH": "*",
        "X-HTTP-Method": "DELETE",
    },
};
var updateHeader = {
    headers: {
        "content-type": "application/json;odata.metadata=full",
        accept: "application/json;odata.metadata=full",
        "X-HTTP-Method": "MERGE",
        "IF-MATCH": "*",
    },
};
var SPService = (function () {
    function SPService(context) {
        var _this = this;
        this.context = context;
        this.webUrl = this.context.pageContext.web.absoluteUrl;
        this.serverUrl = this.context.pageContext.web.serverRelativeUrl;
        this.siteUrl = this.context.pageContext.site.absoluteUrl;
        this.loggedUserName = this.context.pageContext.user.displayName;
        this.loggedUserEmail = this.context.pageContext.user.email;
        this.loggedUserId = this.context.pageContext.legacyPageContext.userId;
        this.adWebUrl = window.location.origin + ":2023/ADExplorer";
        this.getServiceUrl = function (url) { return __awaiter(_this, void 0, void 0, function () {
            var _this = this;
            return __generator(this, function (_a) {
                return [2 /*return*/, this.context.spHttpClient
                        .get(url, SPHttpClient.configurations.v1, {
                        headers: getHeader.headers,
                    })
                        .then(function (response) { return __awaiter(_this, void 0, void 0, function () {
                        var jsonResponse, responseValue, jsonResponse, error, _a;
                        return __generator(this, function (_b) {
                            switch (_b.label) {
                                case 0:
                                    if (!response.ok) return [3 /*break*/, 2];
                                    return [4 /*yield*/, response.json()];
                                case 1:
                                    jsonResponse = _b.sent();
                                    responseValue = {
                                        hasError: false,
                                        value: jsonResponse.value,
                                    };
                                    return [2 /*return*/, responseValue];
                                case 2: return [4 /*yield*/, response.json()];
                                case 3:
                                    jsonResponse = _b.sent();
                                    _a = {
                                        hasError: true
                                    };
                                    return [4 /*yield*/, jsonResponse.error];
                                case 4:
                                    error = (_a.error = _b.sent(),
                                        _a);
                                    return [2 /*return*/, Promise.reject(error)];
                            }
                        });
                    }); })
                        .catch(function (error) {
                        //console.error(error ? error.message : "");
                        console.error(error);
                        return error;
                    })];
            });
        }); };
        this.getDepInfo = function (depName) {
            var queryParams = {
                queryParams: {
                    AllowEmailAddresses: true,
                    AllowMultipleEntities: false,
                    AllUrlZones: false,
                    MaximumEntitySuggestions: 5,
                    PrincipalSource: 15,
                    PrincipalType: 12,
                    QueryString: depName,
                },
            };
            var url = _this.webUrl +
                "/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser";
            var options = {
                headers: postHeader.headers,
                body: queryParams,
            };
            return _this.post(url, options)
                .then(function (result) {
                var resultKey = JSON.parse(result.value.value);
                var ensuruserParam;
                if (resultKey.length > 0) {
                    ensuruserParam = { logonName: resultKey[0].Key };
                }
                return _this.getFinalDepInfo(ensuruserParam);
            })
                .catch(function (err) {
                throw new Error(err);
            });
        };
    }
    SPService.prototype.getListItems = function (listName) {
        return this.context.spHttpClient
            .get(this.webUrl + "/_api/web/lists/getbytitle('" + listName + "')/items?$top=15", SPHttpClient.configurations.v1)
            .then(function (response) { return response.json(); })
            .then(function (data) { return data; }, function (error) { return error; });
    };
    SPService.prototype.changeDateFormat = function (date) {
        var insertedDate = new Date(date);
        var insertedDate2 = this.getFormattedResult(insertedDate.getMonth() + 1) +
            "/" +
            this.getFormattedResult(insertedDate.getDate()) +
            "/" +
            this.getFormattedResult(insertedDate.getFullYear());
        var returnedDate = insertedDate2.split("/");
        return returnedDate[2] + "-" + returnedDate[0] + "-" + returnedDate[1];
    };
    SPService.prototype.getByUrl = function (url) {
        return this.get(url, false)
            .then(function (result) {
            return result;
        })
            .catch(function (err) { return err; });
    };
    SPService.prototype.getAllItems = function (listName) {
        var url = this.webUrl + "/_api/web/lists/getByTitle('" + listName + "')/items";
        return this.get(url).then(function (result) {
            return result;
        });
    };
    SPService.prototype.getFormattedResult = function (num) {
        if (num <= 9) {
            return "0" + num;
        }
        return num;
    };
    SPService.prototype.getItemById = function (listName, id) {
        var url = this.webUrl +
            "/_api/web/lists/getByTitle('" +
            listName +
            "')/items?$select=*,EncodedAbsUrl,FileLeafRef&$filter=Id eq " +
            id;
        // return this.get(url, false)
        return this.getServiceUrl(url)
            .then(function (result) {
            return result;
        })
            .catch(function (err) { return err; });
    };
    SPService.prototype.getFilteredItems = function (listName, query) {
        var url = this.webUrl +
            "/_api/web/lists/getByTitle('" +
            listName +
            "')/items" +
            query;
        return this.get(url, false)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.getFieldsChoices = function (listName, fieldName) {
        var url = this.webUrl +
            "/_api/web/lists/getByTitle('" +
            listName +
            "')/fields/getByTitle('" +
            fieldName +
            "')/Choices";
        return this.get(url, false)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.editAndGet = function (listName, id, inputs) {
        var _this = this;
        return this.updateItem(listName, inputs, id).then(function (response) {
            return _this.getItemById(listName, id).then(function (json) {
                return json;
            });
        });
    };
    SPService.prototype.getAllFiles = function (serverRelativeUrl) {
        var restUrl = this.webUrl + "/_api/web/getFolderByServerRelativeUrl('" + serverRelativeUrl + "')/files";
        return this.get(restUrl, true).then(function (json) {
            return json;
        });
    };
    SPService.prototype.getFilteredFiles = function (serverRelativeUrl, query) {
        var restUrl = this.webUrl + "/_api/web/getFolderByServerRelativeUrl('" + serverRelativeUrl + "')/files" + query;
        return this.get(restUrl, true).then(function (json) {
            return json;
        });
    };
    SPService.prototype.getAllFolders = function (serverRelativeUrl) {
        var restUrl = this.webUrl + "/_api/web/getFolderByServerRelativeUrl('" + serverRelativeUrl + "')/folders";
        return this.get(restUrl, true).then(function (json) {
            return json;
        });
    };
    SPService.prototype.getFilteredFolder = function (serverRelativeUrl, query) {
        var restUrl = this.webUrl + "/_api/web/getFolderByServerRelativeUrl('" + serverRelativeUrl + "')/folders" + query;
        return this.get(restUrl, true).then(function (json) {
            return json;
        });
    };
    SPService.prototype.getLibraryInformationByName = function (libraryName) {
        var restUrl = this.webUrl + "/_api/web/folders?$filter=Name eq '" + libraryName + "'";
        return this.get(restUrl)
            .then(function (json) {
            return json;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.getInformationUsingServerRelativeUrl = function (serverRelativeUrl) {
        var restUrl = this.webUrl + "/_api/web/getFolderByServerRelativeUrl('" + serverRelativeUrl + "')";
        return this.get(restUrl)
            .then(function (json) {
            return json;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.getItemCount = function (listName) {
        var restUrl = this.webUrl + "/_api/web/lists/getByTitle('" + listName + "')/ItemCount";
        return this.get(restUrl)
            .then(function (json) {
            return json;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.get = function (url, check) {
        var _this = this;
        if (check === void 0) { check = false; }
        return this.context.spHttpClient
            .get(url, SPHttpClient.configurations.v1, {
            headers: getHeader.headers,
        })
            .then(function (response) { return __awaiter(_this, void 0, void 0, function () {
            return __generator(this, function (_a) {
                return [2 /*return*/, response.json().then(function (json) {
                        return json;
                    })];
            });
        }); })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.postItem = function (listName, data) {
        var url = this.webUrl + "/_api/web/lists/getByTitle('" + listName + "')/items";
        var options = {
            headers: postHeader.headers,
            body: data,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            throw err;
        });
    };
    SPService.prototype.createList = function (listName, description) {
        if (description === void 0) { description = ""; }
        var url = this.webUrl + "/_api/web/lists";
        var options = {
            headers: postHeader.headers,
            body: {
                // '__metadata': { 'type': 'SP.List' },
                BaseTemplate: 100,
                Title: listName,
                Description: description,
            },
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            throw err;
        });
    };
    SPService.prototype.createSiteField = function (fieldName, groupName) {
        var url = this.webUrl + "/_api/web/fields";
        var options = {
            headers: postHeader.headers,
            body: {
                // '__metadata': { 'type': 'SP.Field' },
                Title: fieldName,
                FieldTypeKind: 2,
                Group: groupName,
            },
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            throw err;
        });
    };
    SPService.prototype.createFieldsForAList = function (listName, fieldsDefinition) {
        var _this = this;
        Promise.all(fieldsDefinition.map(function (fieldDefinition) {
            _this.createFieldForAList(listName, fieldDefinition);
        })).then(function () {
            return;
        });
    };
    SPService.prototype.createFieldForAList = function (listName, fieldDefinition) {
        var url = this.webUrl + ("/_api/web/lists/getByTitle('" + listName + "')/fields");
        var options = {
            headers: postHeader.headers,
            body: fieldDefinition,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            throw err;
        });
    };
    SPService.prototype.postNotification = function (listName, data) {
        var url = window.location.origin + "/sites/portal/" +
            "/_api/web/lists/getByTitle('" +
            listName +
            "')/items";
        var options = {
            headers: postHeader.headers,
            body: data,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            throw err;
        });
    };
    SPService.prototype.updateItem = function (listName, data, id, toJson) {
        if (toJson === void 0) { toJson = true; }
        var url = this.webUrl +
            "/_api/web/lists/getByTitle('" +
            listName +
            "')/items(" +
            id +
            ")";
        var options = {
            headers: updateHeader.headers,
            body: data,
        };
        return this.post(url, options, toJson)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            throw err;
        });
    };
    SPService.prototype.updateFileMetaData = function (fileServerRelativeUrl, data) {
        var url = this.webUrl +
            ("/_api/web/getFileByServerRelativeUrl('" + fileServerRelativeUrl + "')/ListItemAllFields");
        var options = {
            headers: updateHeader.headers,
            body: data,
        };
        return this.post(url, options, true)
            .then(function (result) {
            return result;
        })
            .catch(function (err) { return err; });
    };
    SPService.prototype.updateFolderMetaData = function (folderServerRelativeUrl, data) {
        var url = this.webUrl +
            ("/_api/web/getFolderByServerRelativeUrl('" + folderServerRelativeUrl + "')/ListItemAllFields");
        var options = {
            headers: updateHeader.headers,
            body: data,
        };
        return this.post(url, options, true)
            .then(function (result) {
            return result;
        })
            .catch(function (err) { return err; });
    };
    SPService.prototype.postFile = function (listName, file) {
        var url = this.webUrl +
            "/_api/web/lists/getByTitle('" +
            listName +
            "')/RootFolder/files/add(url='" +
            file.name +
            "',overwrite=true)?$expand=ListItemAllFields";
        var options = {
            headers: postHeader.headers,
            body: file,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            return err;
        });
    };
    // uploadFile(serverRelativeUrl, file): Promise<any> {
    //     sp.web.getFolderByServerRelativeUrl( this.webUrl +
    //       "/_api/web/getFolderByServerRelativeUrl('" +
    //       serverRelativeUrl).files.add(file.name, file, true).then(f => {
    //
    //       f.file.getItem().then(item => {
    //         item.update({
    //           Title: "Metadata Updated"
    //         }).then((result) => {
    //
    //
    //
    //             return result;
    //
    //         }) .catch((err) => {
    //
    //           Promise.reject(err);
    //           return err;
    //         });
    //       }).catch((err) => {
    //
    //         Promise.reject(err);
    //         return err;
    //       });
    //     }).catch((err) => {
    //
    //       Promise.reject(err);
    //       return err;
    //     });
    //   }
    SPService.prototype.postFileByServerRelativeUrl = function (serverRelativeUrl, file) {
        var url = this.webUrl +
            "/_api/web/getFolderByServerRelativeUrl('" +
            serverRelativeUrl +
            "')/files/add(url='" +
            file.name +
            "',overwrite=true)?$expand=ListItemAllFields";
        var options = {
            headers: postHeader.headers,
            body: file,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.createFile = function (listName, fileName) {
        var url = this.webUrl +
            "/_api/web/GetFolderByServerRelativeUrl('" +
            this.serverUrl +
            "/" +
            listName +
            "')/files/add(url='" +
            fileName +
            "',overwrite=true)?$expand=ListItemAllFields";
        var options = {
            headers: postHeader.headers,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.createFileByServerRelativeUrl = function (folderServerRelativeUrl, fileName) {
        var url = this.webUrl +
            "/_api/web/GetFolderByServerRelativeUrl('" +
            folderServerRelativeUrl +
            "')/files/add(url='" +
            fileName +
            "',overwrite=true)?$expand=ListItemAllFields";
        var options = {
            headers: postHeader.headers,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.createFolder = function (serverRelativeUrl, folderName) {
        var url = this.webUrl + "/_api/web/getFolderByServerRelativeUrl('" + serverRelativeUrl + "')/folders/add(url='" + folderName + "')";
        var options = {
            headers: postHeader.headers,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            return err;
        });
    };
    SPService.prototype.moveFile = function (listName, originalFileName, newFileName) {
        var url = this.webUrl +
            "/_api/web/getfilebyserverrelativeurl('" +
            this.serverUrl +
            "/" +
            listName +
            "/" +
            originalFileName +
            "')/moveto(newurl = '" +
            this.serverUrl +
            "/" +
            listName +
            "/" +
            newFileName +
            "', flags = 1)";
        var options = {
            headers: updateHeader.headers,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) { return err; });
    };
    SPService.prototype.moveFolder = function (listName, originalFileName, newFileName) {
        var url = this.webUrl +
            "/_api/web/getfilebyserverrelativeurl('" +
            this.serverUrl +
            "/" +
            listName +
            "/" +
            originalFileName +
            "')/moveto(newurl = '" +
            this.serverUrl +
            "/" +
            listName +
            "/" +
            newFileName +
            "', flags = 1)";
        var options = {
            headers: updateHeader.headers,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) { return err; });
    };
    SPService.prototype.deleteItem = function (listName, id) {
        var url = this.webUrl +
            "/_api/web/lists/getByTitle('" +
            listName +
            "')/items(" +
            id +
            ")";
        var options = {
            headers: deleteHeader.headers,
        };
        return this.context.spHttpClient
            .post(url, SPHttpClient.configurations.v1, options)
            .then(function (response) {
            return response.json();
        })
            .catch(function (err) {
            return err;
        });
    };
    // async post(url: string, postInformation, check = false): Promise<any> {
    //   if (check) {
    //     return await this.dateConverter
    //       .toEuropean(postInformation.body)
    //       .then((response) => {
    //         const options: ISPHttpClientOptions = {
    //           headers: postInformation.headers,
    //           body: JSON.stringify(response),
    //         };
    //         return this.context.spHttpClient
    //           .post(url, SPHttpClient.configurations.v1, options)
    //           .then((result) => {
    //             return result.json().then((json) => {
    //               return json;
    //             });
    //           })
    //           .catch((err) => {
    //
    //             Promise.reject(err);
    //           });
    //       });
    //   }
    SPService.prototype.post = function (url, postInformation, toJson) {
        if (toJson === void 0) { toJson = true; }
        return __awaiter(this, void 0, void 0, function () {
            var _this = this;
            var options;
            return __generator(this, function (_a) {
                options = {
                    headers: postInformation.headers,
                    body: JSON.stringify(postInformation.body),
                };
                return [2 /*return*/, this.context.spHttpClient
                        .post(url, SPHttpClient.configurations.v1, options)
                        .then(function (response) { return __awaiter(_this, void 0, void 0, function () {
                        var jsonResponse, responseValue, jsonResponse, error, _a;
                        return __generator(this, function (_b) {
                            switch (_b.label) {
                                case 0:
                                    if (!toJson) return [3 /*break*/, 6];
                                    if (!response.ok) return [3 /*break*/, 2];
                                    return [4 /*yield*/, response.json()];
                                case 1:
                                    jsonResponse = _b.sent();
                                    responseValue = {
                                        hasError: false,
                                        value: jsonResponse,
                                    };
                                    return [2 /*return*/, responseValue];
                                case 2: return [4 /*yield*/, response.json()];
                                case 3:
                                    jsonResponse = _b.sent();
                                    _a = {
                                        hasError: true
                                    };
                                    return [4 /*yield*/, jsonResponse.error];
                                case 4:
                                    error = (_a.error = _b.sent(),
                                        _a);
                                    return [2 /*return*/, Promise.reject(error)];
                                case 5: return [3 /*break*/, 7];
                                case 6: return [2 /*return*/, response];
                                case 7: return [2 /*return*/];
                            }
                        });
                    }); })
                        .catch(function (error) {
                        //console.error(error ? error.message : "");
                        console.error(error);
                        return error;
                    })];
            });
        });
    };
    SPService.prototype.isCurrentUserInGroup = function (groupName) {
        return __awaiter(this, void 0, void 0, function () {
            var url;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        url = this.context.pageContext.web.absoluteUrl +
                            "/_api/Web/SiteGroups/GetByName('" +
                            groupName +
                            "')/Users?$filter=email eq '" +
                            this.loggedUserEmail +
                            "'";
                        return [4 /*yield*/, this.get(url).then(function (response) {
                                var result = false;
                                if (response.length > 0) {
                                    result = true;
                                }
                                else {
                                    result = false;
                                }
                                return result;
                            })];
                    case 1: return [2 /*return*/, _a.sent()];
                }
            });
        });
    };
    SPService.prototype.getListInformationByName = function (listName) {
        var restUrl = this.webUrl + "/_api/web/lists/getByTitle('" + listName + "')";
        return this.get(restUrl).then(function (json) {
            return json;
        });
    };
    SPService.prototype.getAllGroupsOfAUser = function () {
        var restUrl = this.context.pageContext.web.absoluteUrl + "/_api/web/currentuser/?$expand=groups";
        return this.get(restUrl)
            .then(function (json) {
            return json.Groups;
        })
            .catch(function (err) {
            console.error(err);
        });
    };
    SPService.prototype.getDepartmentsFromAD = function (Ou) {
        var depdata = [];
        return fetch(window.location.origin + ":2023/adexplorer/getorgstr?ou=" + Ou)
            .then(function (response) { return response.json(); })
            .then(function (data) {
            depdata = data;
            return depdata;
        })
            .catch(function (error) {
            console.error(error);
        });
    };
    SPService.prototype.getMyProperties = function () {
        var url = this.webUrl + "/_api/SP.UserProfiles.PeopleManager/GetMyProperties";
        return this.get(url, false).then(function (res) {
            return res;
        });
    };
    SPService.prototype.getUserDepartmentFromAD = function (userName) {
        return fetch(window.location.origin + ":2023/ADExplorer/getUserOU/?UserName=" + userName)
            .then(function (data) {
            var depdata = data;
            return depdata;
        })
            .catch(function (error) {
            console.error(error);
        });
    };
    SPService.prototype.getUserSubDepartments = function (userName, siteName) {
        return fetch(this.adWebUrl + "/GetSubOU/?OU=" + siteName + "&Parent=" + userName)
            .then(function (data) {
            var depdata = data;
            return depdata;
        })
            .catch(function (error) {
            console.error(error);
        });
    };
    SPService.prototype.getPermissionIds = function () {
        var restUrl = this.webUrl + "/_api/web/lists/getByTitle('Permissions')/items?&$select=*&$orderby=Created%20desc";
        return this.get(restUrl)
            .then(function (data) {
            return data;
        })
            .catch(function (err) {
            console.error(err);
        });
    };
    SPService.prototype.getUserRole = function (userID) {
        var restUrl = this.webUrl + "/_api/web/lists/getByTitle('UserRole')/items?&$select=*&$orderby=Created%20desc&$filter=UserId eq '" + userID + "'";
        return this.get(restUrl)
            .then(function (data) {
            return data;
        })
            .catch(function (err) {
            console.error(err);
        });
    };
    SPService.prototype.getUserRoleResources = function (roleIds) {
        var values = roleIds;
        var filterConditions = values
            .map(function (value) { return "Role/Id eq '" + value + "'"; })
            .join(" or ");
        var restUrl = this.webUrl + "/_api/web/lists/getByTitle('RoleResource')/items?&$select=*,Role/Id,PageResource/PageCode&$expand=Role,PageResource&$orderby=Created%20desc&$filter=" + filterConditions;
        return this.get(restUrl)
            .then(function (data) {
            return data;
        })
            .catch(function (err) {
            console.error(err);
        });
    };
    SPService.prototype.getUserRolePermissions = function (roleresourceIds) {
        var values = roleresourceIds;
        var filterConditions = values
            .map(function (value) { return "RoleResourceId eq '" + value + "'"; })
            .join(" or ");
        var restUrl = this.webUrl + "/_api/web/lists/getByTitle('RolePermission')/items?$top=200&$select=*,Permission/Id&$expand=Permission&$orderby=Created%20desc&$filter=" + filterConditions;
        return this.get(restUrl)
            .then(function (data) {
            return data;
        })
            .catch(function (err) {
            console.error(err);
        });
    };
    SPService.prototype.getPageCodes = function () {
        var restUrl = this.webUrl + "/_api/web/lists/getByTitle('PageResource')/items?&$select=*&$orderby=Created%20desc";
        return this.get(restUrl)
            .then(function (data) {
            return data;
        })
            .catch(function (err) {
            console.error(err);
        });
    };
    SPService.prototype.getParentSiteDetail = function () {
        var url = this.webUrl + "/_api/site/RootWeb";
        return this.get(url)
            .then(function (response) {
            return response;
        })
            .catch(function (err) {
            throw new Error("error");
        });
    };
    SPService.prototype.getFinalDepInfo = function (data) {
        var url = this.webUrl + "/_api/web/ensureuser";
        var options = {
            headers: postHeader.headers,
            body: data,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            throw new Error(err);
        });
    };
    SPService.prototype.getUserDepartment = function () {
        var url = this.webUrl + "/_api/SP.UserProfiles.PeopleManager/GetUserProfilePropertyFor(accountName=@v,propertyName='Department')?@v='" + this.context.pageContext.user.loginName + "'";
        return this.get(url)
            .then(function (response) {
            return response.value;
        })
            .catch(function (err) {
            throw new Error("error");
        });
    };
    SPService.prototype.createNotification = function (data) {
        var url = this.siteUrl +
            "/_api/web/lists/getByTitle('Notification_associated_task')/items";
        var options = {
            headers: postHeader.headers,
            body: data,
        };
        return this.post(url, options)
            .then(function (result) {
            return result;
        })
            .catch(function (err) {
            throw err;
        });
    };
    return SPService;
}());
export default SPService;

//# sourceMappingURL=SPServices.js.map

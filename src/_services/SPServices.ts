import { WebPartContext } from "@microsoft/sp-webpart-base";
import {
  SPHttpClientResponse,
  SPHttpClient,
  ISPHttpClientOptions,
} from "@microsoft/sp-http";
import { CanvasA11yManager } from "@microsoft/sp-webpart-base/lib/utils/CanvasA11y";
// import reject = Promise.reject;
const getHeader = {
  headers: {
    accept: "application/json;",
  },
};
const postHeader = {
  headers: {
    "content-type": "application/json;odata.metadata=full",
    accept: "application/json;odata.metadata=full",
  },
};
const deleteHeader = {
  headers: {
    "content-type": "application/json;odata.metadata=full",
    "IF-MATCH": "*",
    "X-HTTP-Method": "DELETE",
  },
};
const updateHeader = {
  headers: {
    "content-type": "application/json;odata.metadata=full",
    accept: "application/json;odata.metadata=full",
    "X-HTTP-Method": "MERGE",
    "IF-MATCH": "*",
  },
};

export default class SPService {
  constructor(private context: WebPartContext) {}
  webUrl = this.context.pageContext.web.absoluteUrl;
  serverUrl = this.context.pageContext.web.serverRelativeUrl;
  siteUrl = this.context.pageContext.site.absoluteUrl;

  loggedUserName = this.context.pageContext.user.displayName;
  loggedUserEmail = this.context.pageContext.user.email;
  loggedUserId = this.context.pageContext.legacyPageContext.userId;
  adWebUrl = `${window.location.origin}:2023/ADExplorer`;

  getListItems(listName: string) {
    return this.context.spHttpClient
      .get(
        `${this.webUrl}/_api/web/lists/getbytitle('${listName}')/items?$top=15`,
        SPHttpClient.configurations.v1
      )
      .then((response: SPHttpClientResponse) => response.json())
      .then(
        (data) => data,
        (error) => error
      );
  }

  changeDateFormat(date) {
    const insertedDate = new Date(date);
    let insertedDate2 =
      this.getFormattedResult(insertedDate.getMonth() + 1) +
      "/" +
      this.getFormattedResult(insertedDate.getDate()) +
      "/" +
      this.getFormattedResult(insertedDate.getFullYear());
    let returnedDate = insertedDate2.split("/");
    return returnedDate[2] + "-" + returnedDate[0] + "-" + returnedDate[1];
  }

  getByUrl(url) {
    return this.get(url, false)
      .then((result) => {
        return result;
      })
      .catch((err) => err);
  }

  getAllItems(listName: string): Promise<any> {
    const url: string =
      this.webUrl + "/_api/web/lists/getByTitle('" + listName + "')/items";
    return this.get(url).then((result) => {
      return result;
    });
  }

  getFormattedResult(num) {
    if (num <= 9) {
      return `0${num}`;
    }
    return num;
  }

  getItemById(listName: string, id) {
    const url: string =
      this.webUrl +
      "/_api/web/lists/getByTitle('" +
      listName +
      "')/items?$select=*,EncodedAbsUrl,FileLeafRef&$filter=Id eq " +
      id;
    // return this.get(url, false)
    return this.getServiceUrl(url)
      .then((result) => {
        return result;
      })
      .catch((err) => err);
  }

  getFilteredItems(listName: string, query): Promise<any> {
    const url: string =
      this.webUrl +
      "/_api/web/lists/getByTitle('" +
      listName +
      "')/items" +
      query;

    return this.get(url, false)
      .then((result) => {
        return result;
      })
      .catch((err) => {
        return err;
      });
  }

  getFieldsChoices(listName: string, fieldName: string): Promise<any> {
    const url: string =
      this.webUrl +
      "/_api/web/lists/getByTitle('" +
      listName +
      "')/fields/getByTitle('" +
      fieldName +
      "')/Choices";

    return this.get(url, false)
      .then((result) => {
        return result;
      })
      .catch((err) => {
        return err;
      });
  }

  editAndGet(listName: string, id, inputs): Promise<any> {
    return this.updateItem(listName, inputs, id).then((response) => {
      return this.getItemById(listName, id).then((json) => {
        return json;
      }) as Promise<any>;
    });
  }

  getAllFiles(serverRelativeUrl: string): Promise<any> {
    const restUrl = `${this.webUrl}/_api/web/getFolderByServerRelativeUrl('${serverRelativeUrl}')/files`;
    return this.get(restUrl, true).then((json) => {
      return json;
    });
  }

  getFilteredFiles(serverRelativeUrl: string, query: string): Promise<any> {
    const restUrl = `${this.webUrl}/_api/web/getFolderByServerRelativeUrl('${serverRelativeUrl}')/files${query}`;

    return this.get(restUrl, true).then((json) => {
      return json;
    });
  }

  getAllFolders(serverRelativeUrl: string): Promise<any> {
    const restUrl = `${this.webUrl}/_api/web/getFolderByServerRelativeUrl('${serverRelativeUrl}')/folders`;
    return this.get(restUrl, true).then((json) => {
      return json;
    });
  }

  getFilteredFolder(serverRelativeUrl: string, query: string): Promise<any> {
    const restUrl = `${this.webUrl}/_api/web/getFolderByServerRelativeUrl('${serverRelativeUrl}')/folders${query}`;

    return this.get(restUrl, true).then((json) => {
      return json;
    });
  }

  getLibraryInformationByName(libraryName: string): Promise<any> {
    const restUrl = `${this.webUrl}/_api/web/folders?$filter=Name eq '${libraryName}'`;
    return this.get(restUrl)
      .then((json) => {
        return json;
      })
      .catch((err) => {
        return err;
      });
  }

  getInformationUsingServerRelativeUrl(serverRelativeUrl): Promise<any> {
    const restUrl = `${this.webUrl}/_api/web/getFolderByServerRelativeUrl('${serverRelativeUrl}')`;
    return this.get(restUrl)
      .then((json) => {
        return json;
      })
      .catch((err) => {
        return err;
      });
  }

  getItemCount(listName: string) {
    const restUrl = `${this.webUrl}/_api/web/lists/getByTitle('${listName}')/ItemCount`;
    return this.get(restUrl)
      .then((json) => {
        return json;
      })
      .catch((err) => {
        return err;
      });
  }

  getServiceUrl = async (url: string): Promise<any> => {
    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1, {
        headers: getHeader.headers,
      })
      .then(async (response: any): Promise<any> => {
        if (response.ok) {
          let jsonResponse = await response.json();

          let responseValue = {
            hasError: false,
            value: jsonResponse.value,
          };
          return responseValue;
        } else {
          let jsonResponse = await response.json();
          let error = {
            hasError: true,
            error: await jsonResponse.error,
          };
          return Promise.reject(error);
        }
      })
      .catch((error: any): void => {
        //console.error(error ? error.message : "");
        console.error(error);
        return error;
      });
  };

  get(url: string, check: Boolean = false): Promise<any> {
    return this.context.spHttpClient
      .get(url, SPHttpClient.configurations.v1, {
        headers: getHeader.headers,
      })
      .then(async (response) => {
        return response.json().then((json) => {
          return json;
        });
      })
      .catch((err) => {
        return err;
      });
  }

  postItem(listName: string, data: any): Promise<any> {
    const url: string =
      this.webUrl + "/_api/web/lists/getByTitle('" + listName + "')/items";
    const options: ISPHttpClientOptions = {
      headers: postHeader.headers,
      body: data,
    };
    return this.post(url, options)
      .then((result) => {
        return result;
      })
      .catch((err) => {
        throw err;
      });
  }

  createList(listName: string, description: string = "") {
    const url: string = this.webUrl + "/_api/web/lists";
    const options: any = {
      headers: postHeader.headers,
      body: {
        // '__metadata': { 'type': 'SP.List' },
        BaseTemplate: 100, // 100 represents a custom list, you can change it as per your requirement
        Title: listName,
        Description: description,
      },
    };
    return this.post(url, options)
      .then((result) => {
        return result;
      })
      .catch((err) => {
        throw err;
      });
  }
  createSiteField(fieldName: string, groupName: string) {
    const url: string = this.webUrl + "/_api/web/fields";
    const options: any = {
      headers: postHeader.headers,
      body: {
        // '__metadata': { 'type': 'SP.Field' },
        Title: fieldName,
        FieldTypeKind: 2,
        Group: groupName,
      },
    };
    return this.post(url, options)
      .then((result) => {
        return result;
      })
      .catch((err) => {
        throw err;
      });
  }

  createFieldsForAList(listName: string, fieldsDefinition: any[]) {
    Promise.all(
      fieldsDefinition.map((fieldDefinition) => {
        this.createFieldForAList(listName, fieldDefinition);
      })
    ).then(() => {
      return;
    });
  }

  createFieldForAList(listName: string, fieldDefinition: any) {
    const url: string =
      this.webUrl + `/_api/web/lists/getByTitle('${listName}')/fields`;
    const options: any = {
      headers: postHeader.headers,
      body: fieldDefinition,
    };
    return this.post(url, options)
      .then((result) => {
        return result;
      })
      .catch((err) => {
        throw err;
      });
  }

  postNotification(listName: string, data: any): Promise<any> {
    const url: string =
      `${window.location.origin}/sites/portal/` +
      "/_api/web/lists/getByTitle('" +
      listName +
      "')/items";
    const options: ISPHttpClientOptions = {
      headers: postHeader.headers,
      body: data,
    };
    return this.post(url, options)
      .then((result) => {
        return result;
      })
      .catch((err) => {
        throw err;
      });
  }

  updateItem(listName: string, data: any, id, toJson = true): Promise<any> {
    const url: string =
      this.webUrl +
      "/_api/web/lists/getByTitle('" +
      listName +
      "')/items(" +
      id +
      ")";
    const options: ISPHttpClientOptions = {
      headers: updateHeader.headers,
      body: data,
    };
    return this.post(url, options, toJson)
      .then((result) => {
        return result;
      })
      .catch((err) => {
        throw err;
      });
  }

  updateFileMetaData(fileServerRelativeUrl, data) {
    const url: string =
      this.webUrl +
      `/_api/web/getFileByServerRelativeUrl('${fileServerRelativeUrl}')/ListItemAllFields`;
    const options: ISPHttpClientOptions = {
      headers: updateHeader.headers,
      body: data,
    };
    return this.post(url, options, true)
      .then((result) => {
        return result;
      })
      .catch((err) => err);
  }

  updateFolderMetaData(folderServerRelativeUrl, data) {
    const url: string =
      this.webUrl +
      `/_api/web/getFolderByServerRelativeUrl('${folderServerRelativeUrl}')/ListItemAllFields`;
    const options: ISPHttpClientOptions = {
      headers: updateHeader.headers,
      body: data,
    };
    return this.post(url, options, true)
      .then((result) => {
        return result;
      })
      .catch((err) => err);
  }

  postFile(listName, file): Promise<any> {
    const url: string =
      this.webUrl +
      "/_api/web/lists/getByTitle('" +
      listName +
      "')/RootFolder/files/add(url='" +
      file.name +
      "',overwrite=true)?$expand=ListItemAllFields";

    const options: ISPHttpClientOptions = {
      headers: postHeader.headers,
      body: file,
    };
    return this.post(url, options)
      .then((result) => {
        return result;
      })
      .catch((err) => {
        return err;
      });
  }

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

  postFileByServerRelativeUrl(serverRelativeUrl, file): Promise<any> {
    const url: string =
      this.webUrl +
      "/_api/web/getFolderByServerRelativeUrl('" +
      serverRelativeUrl +
      "')/files/add(url='" +
      file.name +
      "',overwrite=true)?$expand=ListItemAllFields";

    const options: ISPHttpClientOptions = {
      headers: postHeader.headers,
      body: file,
    };
    return this.post(url, options)
      .then((result) => {
        return result;
      })
      .catch((err) => {
        return err;
      });
  }

  createFile(listName, fileName): Promise<any> {
    const url: string =
      this.webUrl +
      "/_api/web/GetFolderByServerRelativeUrl('" +
      this.serverUrl +
      "/" +
      listName +
      "')/files/add(url='" +
      fileName +
      "',overwrite=true)?$expand=ListItemAllFields";
    const options: ISPHttpClientOptions = {
      headers: postHeader.headers,
    };
    return this.post(url, options)
      .then((result) => {
        return result;
      })
      .catch((err) => {
        return err;
      });
  }

  createFileByServerRelativeUrl(
    folderServerRelativeUrl,
    fileName
  ): Promise<any> {
    const url: string =
      this.webUrl +
      "/_api/web/GetFolderByServerRelativeUrl('" +
      folderServerRelativeUrl +
      "')/files/add(url='" +
      fileName +
      "',overwrite=true)?$expand=ListItemAllFields";
    const options: ISPHttpClientOptions = {
      headers: postHeader.headers,
    };
    return this.post(url, options)
      .then((result) => {
        return result;
      })
      .catch((err) => {
        return err;
      });
  }

  createFolder(serverRelativeUrl: string, folderName: string): Promise<any> {
    const url = `${this.webUrl}/_api/web/getFolderByServerRelativeUrl('${serverRelativeUrl}')/folders/add(url='${folderName}')`;
    const options: ISPHttpClientOptions = {
      headers: postHeader.headers,
    };
    return this.post(url, options)
      .then((result) => {
        return result;
      })
      .catch((err) => {
        return err;
      });
  }

  moveFile(listName, originalFileName, newFileName): Promise<any> {
    const url: string =
      this.webUrl +
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
    const options: ISPHttpClientOptions = {
      headers: updateHeader.headers,
    };
    return this.post(url, options)
      .then((result) => {
        return result;
      })
      .catch((err) => err);
  }

  moveFolder(listName, originalFileName, newFileName): Promise<any> {
    const url: string =
      this.webUrl +
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
    const options: ISPHttpClientOptions = {
      headers: updateHeader.headers,
    };
    return this.post(url, options)
      .then((result) => {
        return result;
      })
      .catch((err) => err);
  }

  deleteItem(listName: string, id) {
    const url: string =
      this.webUrl +
      "/_api/web/lists/getByTitle('" +
      listName +
      "')/items(" +
      id +
      ")";
    const options: ISPHttpClientOptions = {
      headers: deleteHeader.headers,
    };
    return this.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, options)
      .then((response) => {
        return response.json();
      })
      .catch((err) => {
        return err;
      });
  }

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

  async post(url: string, postInformation, toJson = true): Promise<any> {
    const options: ISPHttpClientOptions = {
      headers: postInformation.headers,
      body: JSON.stringify(postInformation.body),
    };
    return this.context.spHttpClient
      .post(url, SPHttpClient.configurations.v1, options)
      .then(async (response: any): Promise<any> => {
        if (toJson) {
          if (response.ok) {
            let jsonResponse = await response.json();

            let responseValue = {
              hasError: false,
              value: jsonResponse,
            };

            return responseValue;
          } else {
            let jsonResponse = await response.json();
            let error = {
              hasError: true,
              error: await jsonResponse.error,
            };
            return Promise.reject(error);
          }
        } else {
          return response;
        }
      })
      .catch((error: any): void => {
        //console.error(error ? error.message : "");
        console.error(error);
        return error;
      });
  }
  async isCurrentUserInGroup(groupName: string): Promise<any> {
    var url =
      this.context.pageContext.web.absoluteUrl +
      "/_api/Web/SiteGroups/GetByName('" +
      groupName +
      "')/Users?$filter=email eq '" +
      this.loggedUserEmail +
      "'";
    return await this.get(url).then((response: any) => {
      let result = false;
      if (response.length > 0) {
        result = true;
      } else {
        result = false;
      }
      return result;
    });
  }
  getListInformationByName(listName: string): Promise<any> {
    const restUrl = `${this.webUrl}/_api/web/lists/getByTitle('${listName}')`;

    return this.get(restUrl).then((json) => {
      return json;
    });
  }

  getAllGroupsOfAUser(): Promise<any> {
    const restUrl = `${this.context.pageContext.web.absoluteUrl}/_api/web/currentuser/?$expand=groups`;
    return this.get(restUrl)
      .then((json) => {
        return json.Groups;
      })
      .catch((err) => {
        console.error(err);
      });
  }

  getDepartmentsFromAD(Ou: string) {
    let depdata = [];
    return fetch(`${window.location.origin}:2023/adexplorer/getorgstr?ou=` + Ou)
      .then((response) => response.json())
      .then((data) => {
        depdata = data;
        return depdata;
      })
      .catch((error) => {
        console.error(error);
      });
  }

  getMyProperties() {
    const url = `${this.webUrl}/_api/SP.UserProfiles.PeopleManager/GetMyProperties`;
    return this.get(url, false).then((res) => {
      return res;
    });
  }

  getUserDepartmentFromAD(userName) {
    return fetch(
      `${window.location.origin}:2023/ADExplorer/getUserOU/?UserName=${userName}`
    )
      .then((data) => {
        const depdata = data;
        return depdata;
      })
      .catch((error) => {
        console.error(error);
      });
  }

  getUserSubDepartments(userName, siteName) {
    return fetch(`${this.adWebUrl}/GetSubOU/?OU=${siteName}&Parent=${userName}`)
      .then((data) => {
        const depdata = data;
        return depdata;
      })
      .catch((error) => {
        console.error(error);
      });
  }

  getPermissionIds() {
    const restUrl = `${this.webUrl}/_api/web/lists/getByTitle('Permissions')/items?&$select=*&$orderby=Created%20desc`;
    return this.get(restUrl)
      .then((data) => {
        return data;
      })
      .catch((err) => {
        console.error(err);
      });
  }

  getUserRole(userID) {
    const restUrl = `${this.webUrl}/_api/web/lists/getByTitle('UserRole')/items?&$select=*&$orderby=Created%20desc&$filter=UserId eq '${userID}'`;
    return this.get(restUrl)
      .then((data) => {
        return data;
      })
      .catch((err) => {
        console.error(err);
      });
  }

  getUserRoleResources(roleIds) {
    const values = roleIds;
    const filterConditions = values
      .map((value) => `Role/Id eq '${value}'`)
      .join(" or ");

    const restUrl = `${this.webUrl}/_api/web/lists/getByTitle('RoleResource')/items?&$select=*,Role/Id,PageResource/PageCode&$expand=Role,PageResource&$orderby=Created%20desc&$filter=${filterConditions}`;
    return this.get(restUrl)
      .then((data) => {
        return data;
      })
      .catch((err) => {
        console.error(err);
      });
  }

  getUserRolePermissions(roleresourceIds) {
    const values = roleresourceIds;
    const filterConditions = values
      .map((value) => `RoleResourceId eq '${value}'`)
      .join(" or ");

    const restUrl = `${this.webUrl}/_api/web/lists/getByTitle('RolePermission')/items?$top=200&$select=*,Permission/Id&$expand=Permission&$orderby=Created%20desc&$filter=${filterConditions}`;
    return this.get(restUrl)
      .then((data) => {
        return data;
      })
      .catch((err) => {
        console.error(err);
      });
  }

  getPageCodes() {
    const restUrl = `${this.webUrl}/_api/web/lists/getByTitle('PageResource')/items?&$select=*&$orderby=Created%20desc`;
    return this.get(restUrl)
      .then((data) => {
        return data;
      })
      .catch((err) => {
        console.error(err);
      });
  }

  getParentSiteDetail(): Promise<any> {
    const url = `${this.webUrl}/_api/site/RootWeb`;
    return this.get(url)
      .then((response: any) => {
        return response;
      })
      .catch((err) => {
        throw new Error(`error`);
      });
  }

  getDepInfo = (depName: string): Promise<any> => {
    const queryParams: any = {
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
    const url: string =
      this.webUrl +
      "/_api/SP.UI.ApplicationPages.ClientPeoplePickerWebServiceInterface.clientPeoplePickerSearchUser";
    const options: ISPHttpClientOptions = {
      headers: postHeader.headers,
      body: queryParams,
    };
    return this.post(url, options)
      .then((result) => {
        const resultKey = JSON.parse(result.value.value);

        let ensuruserParam: any;
        if (resultKey.length > 0) {
          ensuruserParam = { logonName: resultKey[0].Key };
        }

        return this.getFinalDepInfo(ensuruserParam);
      })
      .catch((err) => {
        throw new Error(err);
      });
  };

  getFinalDepInfo(data: any) {
    const url: string = this.webUrl + "/_api/web/ensureuser";
    const options: ISPHttpClientOptions = {
      headers: postHeader.headers,
      body: data,
    };
    return this.post(url, options)
      .then((result) => {
        return result;
      })
      .catch((err) => {
        throw new Error(err);
      });
  }

  getUserDepartment(): Promise<any> {
    const url = `${this.webUrl}/_api/SP.UserProfiles.PeopleManager/GetUserProfilePropertyFor(accountName=@v,propertyName='Department')?@v='${this.context.pageContext.user.loginName}'`;
    return this.get(url)
      .then((response: any) => {
        return response.value;
      })
      .catch((err) => {
        throw new Error(`error`);
      });
  }

  createNotification(data): Promise<any> {
    const url: string =
      this.siteUrl +
      "/_api/web/lists/getByTitle('Notification_associated_task')/items";
    const options: ISPHttpClientOptions = {
      headers: postHeader.headers,
      body: data,
    };
    return this.post(url, options)
      .then((result) => {
        return result;
      })
      .catch((err) => {
        throw err;
      });
  }
}

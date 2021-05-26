import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPHttpClient, IHttpClientOptions } from '@microsoft/sp-http';
import { UserData } from './Helpers';
import * as React from 'react';

declare global {
  interface String {
    escapeSpecialChars(): String;
  }
}
String.prototype.escapeSpecialChars = function () {
  return (this)
    .replace(/\\n/g, "\\n")
    .replace(/\\'/g, "\\'")
    .replace(/\"/g, '\\"')
    .replace(/\\&/g, "\\&")
    .replace(/\\r/g, "\\r")
    .replace(/\\t/g, "\\t")
    .replace(/\\b/g, "\\b")
    .replace(/\\f/g, "\\f");
};

export { };

export abstract class Utils {

  public static adjustTimeZone(dateStr: string): string {

    let date = new Date(dateStr);
    let timeOffsetInMS: number = date.getTimezoneOffset() * 60000;
    date.setTime(date.getTime() - timeOffsetInMS);

    return date.toISOString().replace("T", " ").replace("Z", "").replace(".000", "");
  }

  public static getBreadCrumb(fileUrl: string, serverRelativeUrl: string) {
    let folderPath = "";
    let libraryPath = serverRelativeUrl.replace(/ /g, "%20");
    let itemPath = fileUrl;
    let folderArray = itemPath.split(libraryPath)[1].split("/");
    let folderArraySlice = folderArray.slice(0, folderArray.length);
    let len = folderArraySlice.length;

    for (var i = 0; i < folderArraySlice.length; i++) {
      if (i == len - 1) { //the last folder in the array
        folderPath += folderArraySlice[i];
      }
      else {
        folderPath += folderArraySlice[i];
        folderPath += "/"; //add separator when there are more folders
      }
    }

    return folderPath.replace(/%20/g, " ");
  }

  public static generateGuid() {
    return 'xxxxxxxx-xxxx-xxxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, (c) => {
      let r = Math.random() * 16 | 0, v = c == 'x' ? r : (r & 0x3 | 0x8);
      return v.toString(16);
    });
  }

  public static delay(ms: number) {
    return new Promise(resolve => setTimeout(resolve, ms));
  }

  public static async getUserDataFromUserID(context: WebPartContext, userId: number): Promise<UserData> {

    let requestHeaders: Headers = new Headers();
    requestHeaders.append('content-type', 'application/json');
    requestHeaders.append('accept', 'application/json');

    let httpclientoptions: IHttpClientOptions = {
      method: 'GET',
      body: null,
      headers: requestHeaders
    };

    let response = await context.spHttpClient.fetch(context.pageContext.web.absoluteUrl + "/_api/web/getuserbyid(" + userId + ")",
      SPHttpClient.configurations.v1,
      httpclientoptions);

    let responseParsed = await response.json();

    if (responseParsed.Email == "") {
      responseParsed.Email = responseParsed.UserPrincipalName;
    }

    return new UserData(userId, responseParsed.Email, responseParsed.Title);
  }

  public static async getUserDataFromLoginName(context: WebPartContext, loginName: string): Promise<UserData> {

    let requestHeaders: Headers = new Headers();
    requestHeaders.append('content-type', 'application/json');
    requestHeaders.append('accept', 'application/json');

    let httpclientoptions: IHttpClientOptions = {
      method: 'GET',
      body: null,
      headers: requestHeaders
    };

    let response = await context.spHttpClient.fetch(context.pageContext.web.absoluteUrl + "/_api/web/siteusers(@v)?@v='" + loginName + "'",
      SPHttpClient.configurations.v1,
      httpclientoptions);

    let responseParsed = await response.json();

    if (responseParsed.Email == "") {
      responseParsed.Email = responseParsed.UserPrincipalName;
    }

    return new UserData(responseParsed.Id, responseParsed.Email, responseParsed.Title);
  }

  public static drawVersion = (isVersionDrawn:boolean, version:string) => {
    if (!isVersionDrawn) {
      const elem = document.createElement("label");
      elem.innerText = version;
      elem.className = "versionLabel";
      document.getElementsByClassName("ms-Panel-navigation")[0].prepend(elem);

      return true;
    }
  }

  public static GetLoadingRoller(valueToCheck: boolean) {
    return (<div className="lds-roller" style={valueToCheck == true ? { display: 'none' } : { display: 'block' }} >
      <div></div>
      <div></div>
      <div></div>
      <div></div>
      <div></div>
      <div></div>
      <div></div>
      <div></div>
    </div>);
  }
}

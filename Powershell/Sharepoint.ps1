# https://learn.microsoft.com/en-us/powershell/sharepoint/sharepoint-online/connect-sharepoint-online
Connect-SPOService -Url https://contoso-admin.sharepoint.com -Credential admin@contoso.com

# https://emilymancini.com/2020/01/16/editing-a-news-link-in-sharepoint-online/

$(document).ready(function() {
       var siteUrl = _spPageContextInfo.siteServerRelativeUrl;
       var webUrl = _spPageContextInfo.webServerRelativeUrl;
       var isRootWeb = siteUrl == webUrl;
       if(isRootWeb == false){
        window.location.replace(_spPageContextInfo.siteAbsoluteUrl);
       }
 });

<%@ Page language=”C#” Inherits=”Microsoft.SharePoint.WebPartPages.WebPartPage,Microsoft.SharePoint,Version=12.0.0.0,Culture=neutral,PublicKeyToken=71e9bce111e9429c” %><br><br>window.location.href = [http://new_url_you_want_redirect_it](http://new_url_you_want_redirect_it/);

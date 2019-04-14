import * as url from "url";

  /**
   * Returns the absolute URL according to a Web URL and the server relative URL of a folder
   * @param webUrl The full URL of a web
   * @param serverRelativeUrl The server relative URL of a folder
   * @example
   * // returns "https://contoso.sharepoint.com/sites/team1/Lists/MyList"
   * Utils.getAbsoluteUrl("https://contoso.sharepoint.com/sites/team1/", "/sites/team1/Lists/MyList");
   */
  export const getAbsoluteUrl = (webUrl: string, serverRelativeUrl: string) : string => {
    const uri: url.UrlWithStringQuery = url.parse(webUrl);
    const tenantUrl: string = `${uri.protocol}//${uri.hostname}`;
    if (serverRelativeUrl[0] !== '/') {
      serverRelativeUrl = `/${serverRelativeUrl}`;
    }
    return `${tenantUrl}${serverRelativeUrl}`;
  };


  export const download = (filename: string, content: string) => {

      var element = document.createElement('a');
      element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(content));
      element.setAttribute('download', filename);
  
      element.style.display = 'none';
      document.body.appendChild(element);
  
      element.click();
  
      document.body.removeChild(element);
  };
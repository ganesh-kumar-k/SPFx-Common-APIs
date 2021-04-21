/*! *****************************************************************************
Author : Ganesh Kumar
EMail  : kganeshkumar996@gmail.com
***************************************************************************** */

import { SPHttpClient, SPHttpClientResponse, IHttpClientOptions, IDigestCache, DigestCache } from '@microsoft/sp-http';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as $ from "jquery";

export default class APIHelper {

    private context: WebPartContext;
    private digest: string;

     /**
     * Service constructor
     */
     constructor(_pageContext: WebPartContext) {
        this.context = _pageContext;

        //To get the digest value for AJAX post call
        const digestCache: IDigestCache = _pageContext.serviceScope.consume(DigestCache.serviceKey);
        digestCache.fetchDigest(_pageContext.pageContext.web.serverRelativeUrl).then((digest) => {
            this.digest = digest;
        });
    }

    /*=====================================================
            Uploaded Document using  Rest API
    =======================================================*/
    /**
     * You can upload files up to 2 GB with the REST API.
     * @param serverRelativeUrl Server Relative Url of the folder or library.
     * @param elementId String that specifies the ID value.
     * @param metadata Metadata for the document (optional).
     * @example 
     * var serverRelativeUrl = "/sites/rootsite/subsite/shared documents", 
     * var elementId = "getFile"
     */
    public async addFileToFolderUsingRestApi(serverRelativeUrl:string,elementId:string,metadata?): Promise<any> {
        
        var _this = this;

        var _metadata = JSON.parse(JSON.stringify(metadata)); 

        // Get the file name from the file input control on the page.
        var fileInput:any = $('#'+elementId);
        var parts = fileInput[0].value.split('\\');
        var fileName = parts[parts.length - 1];

        // Get the local file as an array buffer.
        var arrayBuffer = await this.getFileBuffer(elementId);

        // Construct the endpoint.
        var fileCollectionEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/getfolderbyserverrelativeurl('${serverRelativeUrl}')/files/add(overwrite=true, url='${fileName}')`;

        // Send the request and return the response.
        // This call returns the SharePoint file.
        try{

            const response = await $.ajax({
                url: fileCollectionEndpoint,
                type: "POST",
                data: arrayBuffer,
                processData: false,
                headers: {
                    "accept": "application/json;odata=verbose",
                    "X-RequestDigest": this.digest,
                    "content-length": arrayBuffer.byteLength
                }
            });

            console.log(`${fileName} successfully uploaded in ${serverRelativeUrl}`);

            if(_metadata != undefined){
                if(response.d.hasOwnProperty("ListItemAllFields")){  //To check the uri property

                    let fileListItemUri = response.d.ListItemAllFields.__deferred.uri;
                    const listItem = await _this.getListItemUsingRestApi(fileListItemUri);
                    _metadata["__metadata"] = {'type':`${ listItem.d.__metadata.type }`};
                    const result = await _this.updateListItemUsingRestApi(listItem.d.__metadata.uri,_metadata);

                    return result;
                }else{
                    return response;
                }
            }else{
                return response;
            }

        }catch(err){
            console.error(err);
            return err;
        }

    }

    /*=====================================================
            Uploaded Document using  SPHttpClient
    =======================================================*/
    /**
     * You can upload files up to 2 GB with SPHttpClient.
     * @param serverRelativeUrl Server Relative Url of the folder or library.
     * @param elementId String that specifies the ID value.
     * @param metadata Metadata for the document (optional).
     * @example 
     * var serverRelativeUrl = "/sites/rootsite/subsite/shared documents", 
     * var elementId = "getFile"
     */
    public async addFileToFolder(serverRelativeUrl:string,elementId:string,metadata?): Promise<any> {
        
        var _metadata = JSON.parse(JSON.stringify(metadata)); 

        // Get the file name from the file input control on the page.
        var fileInput:any = $('#'+elementId);
        var parts = fileInput[0].value.split('\\');
        var fileName = parts[parts.length - 1];

        // Construct the endpoint.
        var fileCollectionEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/getfolderbyserverrelativeurl('${serverRelativeUrl}')/files/add(overwrite=true, url='${fileName}')`;

        // Construct options
        const header = {  
            "accept": "application/json",
            'Content-type': 'application/json'
        };

        const httpClientOptions: IHttpClientOptions = {  
            body: fileInput[0].files[0],  
            headers: header
        };

        // Send the request and return the response.
        // This call returns the SharePoint file.
        return this.context.spHttpClient.post(fileCollectionEndpoint,SPHttpClient.configurations.v1,httpClientOptions).then((res) => {
            return res.json().then(async (response) => {

                console.log(`${fileName} successfully uploaded in ${serverRelativeUrl}`);

                if(_metadata != undefined){
                    if(response.hasOwnProperty("@odata.id")){  //To check the uri property

                        let fileListItemUri = response["@odata.id"] + "/ListItemAllFields";
                        const listItem = await this.getListItemUsingRestApi(fileListItemUri);
                        const result = await this.updateListItem(listItem.d.__metadata.uri,_metadata);

                        return result;
                    }else{
                        return response;
                    }
                }else{
                    return response;
                }
            },(error) => {
                console.error(error);
                return error;
            });
        });
    }

    /*=====================================================
            Update List Item using  Rest API
    =======================================================*/
    /**
     * To update the list item with the REST API.
     * @param itemUrl URI of the item to update.
     * @param metadata Metadata for the item.
     * @example 
     * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items(1)", 
     */
    public async updateListItemUsingRestApi(itemUrl: string,metadata) : Promise<any> {

        var url = itemUrl;
        var _metadata = JSON.parse(JSON.stringify(metadata));

        //To check and add the list type name if not exists
        if(!_metadata.hasOwnProperty("__metadata")){
            var itemMetadata = await this.getListItemUsingRestApi(itemUrl);
            _metadata["__metadata"] = itemMetadata.d.__metadata;
        }

        var body = JSON.stringify(_metadata);

        // Send the request and return the promise.
        // This call does not return response content from the server.
        const response = await $.ajax({
            url: url,
            type: "POST",
            data: body,
            headers: {
                "accept": "application/json;odata=verbose",  
                "X-RequestDigest": this.digest,  
                "content-Type": "application/json;odata=verbose",  
                "IF-MATCH": "*",  
                "X-HTTP-Method": "MERGE"
            }
        });

        return response;

    }

    /*=====================================================
            Update List Item using  SPHTTPClient
    =======================================================*/
    /**
     * To update the list item with the SPHTTPClient.
     * @param itemUrl URI of the item to update.
     * @param metadata Metadata for the item.
     * @example 
     * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items(1)", 
     */
    public updateListItem(itemUrl:string,metadata) : Promise<any> {

        var url = itemUrl;
        var _metadata = JSON.parse(JSON.stringify(metadata));

        //To remove the __metadata property if exists
        if(_metadata.hasOwnProperty("__metadata")){
            delete _metadata.__metadata;
        }
    
        const header = {  
            'Accept': 'application/json;odata=nometadata',  
            'Content-type': 'application/json;odata=nometadata',  
            'odata-version': '',  
            'IF-MATCH': '*',  
            'X-HTTP-Method': 'MERGE'
        };

        const httpClientOptions: IHttpClientOptions = {  
            headers: header,
            body: JSON.stringify(_metadata)
        };

        // Send the request and return the promise.
        // This call does not return response content from the server.
        return this.context.spHttpClient.post(url,SPHttpClient.configurations.v1,httpClientOptions).then((response: SPHttpClientResponse) => {
            console.log(`Item successfully updated`);
            return response;
        },(error: any) => {
            console.error(error);
            return error;
        });
    }

    /*=====================================================
            Retrieve List Item using  Rest API
    =======================================================*/
    /**
     * To get the list item with the REST API..
     * @param itemUrl URI of the item to update.
     * @example 
     * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items", 
     */
    public async getListItemUsingRestApi(ItemUrl:string) : Promise<any> {
        // Send the request and return the response.
        const response = await $.ajax({
            url: ItemUrl,
            type: "GET",
            headers: { "accept": "application/json;odata=verbose" }
        });

        return response;
    }

    /*=====================================================
            Retrieve the file array buffer
    =======================================================*/
    /**
     * Get the local file as an array buffer.
     */
    private getFileBuffer(fileElementId: string) {
        var fileInput:any = $('#' + fileElementId);
        var deferred = $.Deferred();
        var reader = new FileReader();
        reader.onloadend = function (e:any) {
        deferred.resolve(e.target.result);
        }
        reader.onerror = function (e:any) {
        deferred.reject(e.target.error);
        }
        reader.readAsArrayBuffer(fileInput[0].files[0]);
        return deferred.promise();
    }

}
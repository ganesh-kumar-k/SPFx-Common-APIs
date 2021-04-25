/*! *****************************************************************************
Author : Ganesh Kumar
EMail  : kganeshkumar996@gmail.com
***************************************************************************** */

import { SPHttpClient, SPHttpClientResponse, IHttpClientOptions, IDigestCache, DigestCache } from '@microsoft/sp-http';
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { getFileBuffer } from './CommonFunction';
import * as $ from "jquery";
export default class APIHelper {

    private context: WebPartContext;
    private digest: string;

     /**
     * Service constructor
     * @inheritdoc
     * import APIHelper from '../../../utilities/APIHelper';
     * @example
     * const ApiHelper: APIHelper = new APIHelper(this.props.context);
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
    public async uploadFileToFolderUsingRestApi(serverRelativeUrl: string,elementId: string,metadata?: object): Promise<any> {
        
        var _this = this;

        var fileCount = fileInput[0].files.length;
        if(fileCount == 0)
            return "File is empty";

        var _metadata = JSON.parse(JSON.stringify(metadata)); 

        // Get the file name from the file input control on the page.
        var fileInput:any = $('#'+elementId);
        var parts = fileInput[0].value.split('\\');
        var fileName = parts[parts.length - 1];

        // Get the local file as an array buffer.
        var arrayBuffer = await getFileBuffer(elementId);

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
                    const listItem = await _this.getListItemsUsingRestApi(fileListItemUri);
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
     * @param fileInput File input value
     * @example 
     * var serverRelativeUrl = "/sites/rootsite/subsite/shared documents", 
     * var elementId = "getFile"
     */
    public async uploadFileToFolder(serverRelativeUrl: string,elementId: string,metadata?: object,file?: File): Promise<any> {
        
        // Get the file name from the file input control on the page.
        var fileInput: HTMLInputElement = (<HTMLInputElement>document.getElementById(elementId));

        if(fileInput.files.length == 0)
            return "File is empty";
        
        var _metadata = JSON.parse(JSON.stringify(metadata));

        var fileName = file != undefined ? file.name : fileInput.files.item(0).name;

        var _file = file != undefined ? file : fileInput.files.item(0);

        // Construct the endpoint.
        var fileCollectionEndpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/getfolderbyserverrelativeurl('${serverRelativeUrl}')/files/add(overwrite=true, url='${fileName}')`;

        // Construct headers
        const header = {  
            "accept": "application/json",
            'Content-type': 'application/json'
        };

        const httpClientOptions: IHttpClientOptions = {  
            body: _file,  
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
                        const listItem = await this.getListItemsUsingRestApi(fileListItemUri);
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
    public async updateListItemUsingRestApi(itemUrl: string,metadata: object) : Promise<any> {

        var url = itemUrl;
        var _metadata = JSON.parse(JSON.stringify(metadata));

        //To check and add the list type name if not exists
        if(!_metadata.hasOwnProperty("__metadata")){
            var itemMetadata = await this.getListItemsUsingRestApi(itemUrl);
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
    public updateListItem(itemUrl: string,metadata: object) : Promise<any> {

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
     * @param itemUrl URI of the item to retrieve.
     * @example 
     * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items", 
     */
    public async getListItemsUsingRestApi(ItemUrl: string) : Promise<any> {
        // Send the request and return the response.
        const response = await $.ajax({
            url: ItemUrl,
            type: "GET",
            headers: { "accept": "application/json;odata=verbose" }
        });

        return response;
    }

    /*=====================================================
            Retrieve List Item using  SPHTTPClient
    =======================================================*/
    /**
     * To retrieve the list item with the SPHTTPClient.
     * @param itemUrl URI of the item to retrieve.
     * @example 
     * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items", 
     */
     public getListItems(itemUrl: string) : Promise<any> {
    
        const header = {  
            'Accept': 'application/json;odata=nometadata',  
            'odata-version': ''  
        };

        const httpClientOptions: IHttpClientOptions = {  
            headers: header
        };

        // Send the request and return the promise.
        // This call does not return response content from the server.
        return this.context.spHttpClient.post(itemUrl,SPHttpClient.configurations.v1,httpClientOptions).then((response: SPHttpClientResponse) => {
            console.log(`Item successfully received`);
            return response;
        },(error: any) => {
            console.error(error);
            return error;
        });
    }

    /*=====================================================
            Add List Item using  Rest API
    =======================================================*/
    /**
     * To add the list item with the REST API.
     * @param itemUrl URI of the item to add.
     * @param metadata Metadata for the item.
     * @example 
     * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items", 
     */
     public async addListItemUsingRestApi(itemUrl: string,metadata: object) : Promise<any> {

        var url = itemUrl;
        var _metadata = JSON.parse(JSON.stringify(metadata));

        //To check and add the list type name if not exists
        if(!_metadata.hasOwnProperty("__metadata")){
            var itemMetadata = await this.getListItemsUsingRestApi(itemUrl);
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
                "Accept": "application/json;odata=verbose",
                "Content-Type": "application/json;odata=verbose",
                "X-RequestDigest": this.digest
            }
        });

        return response;

    }

    /*=====================================================
            Update List Item using  SPHTTPClient
    =======================================================*/
    /**
     * To add the list item with the SPHTTPClient.
     * @param itemUrl URI of the item to add.
     * @param metadata Metadata for the item.
     * @example 
     * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items", 
     */
    public addListItem(itemUrl: string,metadata: object) : Promise<any> {

        var url = itemUrl;
        var _metadata = JSON.parse(JSON.stringify(metadata));

        //To remove the __metadata property if exists
        if(_metadata.hasOwnProperty("__metadata")){
            delete _metadata.__metadata;
        }
    
        const header = {  
            'Accept': 'application/json;odata=nometadata',  
            'Content-type': 'application/json;odata=nometadata',  
            'odata-version': '' 
        };

        const httpClientOptions: IHttpClientOptions = {  
            headers: header,
            body: JSON.stringify(_metadata)
        };

        // Send the request and return the promise.
        // This call does not return response content from the server.
        return this.context.spHttpClient.post(url,SPHttpClient.configurations.v1,httpClientOptions).then((response: SPHttpClientResponse) => {
            console.log(`Item successfully added`);
            return response;
        },(error: any) => {
            console.error(error);
            return error;
        });
    }

    /*=====================================================
            Delete List Item using  Rest API
    =======================================================*/
    /**
     * To delete the list item with the REST API.
     * @param itemUrl URI of the item to delete.
     * @example 
     * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items(1)", 
     */
     public async deleteListItemUsingRestApi(itemUrl: string) : Promise<any> {

        var url = itemUrl;

        // Send the request and return the promise.
        // This call does not return response content from the server.
        const response = await $.ajax({
            url: url,
            type: "POST",
            headers: {
                // Accept header: Specifies the format for response data from the server.
                "Accept": "application/json;odata=verbose",
                //Content-Type header: Specifies the format of the data that the client is sending to the server
                "Content-Type": "application/json;odata=verbose",
                // IF-MATCH header: Provides a way to verify that the object being changed has not been changed since it was last retrieved.
                // "IF-MATCH":"*", will overwrite any modification in the object, since it was last retrieved.
                "IF-MATCH": "*",
                //X-HTTP-Method:  The MERGE method updates only the properties of the entity , while the PUT method replaces the existing entity with a new one that you supply in the body of the POST
                "X-HTTP-Method": "DELETE",
                // X-RequestDigest header: When you send a POST request, it must include the form digest value in X-RequestDigest header
                "X-RequestDigest": this.digest
            }
        });

        return response;

    }

    /*=====================================================
            Delete List Item using  SPHTTPClient
    =======================================================*/
    /**
     * To delete the list item with the SPHTTPClient.
     * @param itemUrl URI of the item to delete.
     * @example 
     * var itemUrl = "https://contoso.sharepoint.com/sites/rootsite/subsite/_api/web/lists/getbytitle('Documents')/items(1)", 
     */
    public deleteListItem(itemUrl: string) : Promise<any> {

        var url = itemUrl;
    
        const header = {  
            'Accept': 'application/json;odata=nometadata',  
            'Content-type': 'application/json;odata=verbose',  
            'odata-version': '',  
            'IF-MATCH': '*',  
            'X-HTTP-Method': 'DELETE' 
        };

        const httpClientOptions: IHttpClientOptions = {  
            headers: header,
        };

        // Send the request and return the promise.
        // This call does not return response content from the server.
        return this.context.spHttpClient.post(url,SPHttpClient.configurations.v1,httpClientOptions).then((response: SPHttpClientResponse) => {
            console.log(`Item successfully deleted`);
            return response;
        },(error: any) => {
            console.error(error);
            return error;
        });
    }

    /*=====================================================
            Upload multiple files using SPHTTPClient
    =======================================================*/
    /**
     * You can upload multiple files with the SPHTTPClient.
     * @param serverRelativeUrl Server Relative Url of the folder or library.
     * @param elementId String that specifies the ID value.
     * @param metadata Metadata for the document (optional).
     * @example 
     * var serverRelativeUrl = "/sites/rootsite/subsite/shared documents", 
     * var elementId = "getFile"
     */
    public async uploadMultipleFilesToFolder(serverRelativeUrl: string,elementId: string,metadata?: object): Promise<any> {
       
        // Get values from the file input and text input page controls.
        var fileInput: HTMLInputElement = (<HTMLInputElement>document.getElementById(elementId));

        if(fileInput.files.length == 0)
            return "File is empty";

        var fileCount  = fileInput.files.length;
        var count: number = 0;
        var filesResponse = Array.prototype.map.call(fileInput.files,async (file: File) =>{

            const response = await this.uploadFileToFolder(serverRelativeUrl,elementId,metadata,file);
            count++;
            console.log("Total file uploaded: " + count + " of " + fileCount);
            return response;

        });
        
        const result = await Promise.all(filesResponse);
        return result;
    }

}
import * as React from 'react';
import styles from './DocumentUpload.module.scss';
import { IDocumentUploadProps } from './IDocumentUploadProps';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration,IHttpClientOptions } from '@microsoft/sp-http';
import APIHelper from '../../../utilities/APIHelper';
import * as $ from "jquery"; 
import 'bootstrap/dist/css/bootstrap.min.css';

export default class DocumentUpload extends React.Component<IDocumentUploadProps, {}> {

  private ApiHelper: APIHelper;
  
  constructor(props){
    super(props);
    this.ApiHelper = new APIHelper(this.props.context);
  }

  componentDidMount = () => {
    //this.updatePic();'
    //module.addFileToFolder(this.props.context,)
    let metadata = {};
    this.ApiHelper.addListItemUsingRestApi("",metadata);
    document.getElementById("btn_submit").addEventListener('click',this.addFileToFolder.bind(this));
  }

  // Add the file to the file collection in the Shared Documents folder.
  public async addFileToFolder() {
    // Get the file name from the file input control on the page.

    let fileId = "getFile";
    let serverRelativeUrl = "/sites/practice/john/shared documents";
    let seconds = new Date().getSeconds();
    let metadata = { Title : new Date().getSeconds().toString()};
    let metadata2 = { "__metadata" : "dta",Title : new Date().getSeconds().toString()};


    const result = await this.ApiHelper.addFileToFolderUsingRestApi(serverRelativeUrl,fileId,metadata);
    //const result1 = await this.ApiHelper.addFileToFolder(serverRelativeUrl,fileId,metadata2);
    var url1 = "https://qantler.sharepoint.com/sites/practice/john/_api/web/lists/getbytitle('SpfxPic')/items(28)"
    let metadata1 = { 
      "AdduserId" : [9]
    };
    const result2 = await this.ApiHelper.updateListItem(url1,metadata1);
    alert("Completed");
    // console.log(result);
    // console.log(result1);

  }

  public updatePic() {
    
    let metadata = {};
    //metadata['__metadata'] = { 'type': 'SP.Data.SpfxPicItem' }
    //metadata["UsersStringId"] =  ["9","141"]
    metadata["UsersId"] = [9]


    let url = this.props.context.pageContext.web.absoluteUrl +  "/_api/web/lists/getbytitle('SpfxPic')/items(6)";
    
    const header = {  
      'Accept': 'application/json;odata=nometadata',  
      'Content-type': 'application/json;odata=nometadata',  
      'odata-version': ''  ,
      "IF-MATCH": "*",
      "X-HTTP-Method": "PATCH"
    };

    const httpClientOptions: IHttpClientOptions = {  
      body: JSON.stringify(metadata),  
      headers: header  
    };

    this.props.context.spHttpClient.post(url,SPHttpClient.configurations.v1,httpClientOptions).then((res) => {
      res.json().then(response => {
        alert("Item updated Successfully");
      },error => {
        console.error(error);
      })
    })

  }

  public render(): React.ReactElement<IDocumentUploadProps> {
    return (
      <div className="container-fluid">
          <div>
            <div className="form-group">
              <label htmlFor="userid" className="col-3 col-form-label">User ID</label>
              <div className="col-10">
                <input type="text" className="form-control" id="userid" placeholder="User ID" />
              </div>
            </div>
            <div className="form-group">
              <label htmlFor="dob" className="col-3 col-form-label">Users</label>
              <div className="col-10">
                <input id="getFile" type="file"/>
              </div>
            </div>
            <button type="button" className="btn btn-primary m-3 float-left" id="btn_submit">Upload</button>
          </div>
        </div>
    );
  }
}

import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { escape } from '@microsoft/sp-lodash-subset';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration,IHttpClientOptions } from '@microsoft/sp-http';

import styles from './NojavascriptWebPart.module.scss';
import * as strings from 'NojavascriptWebPartStrings';

import * as $ from "jquery"; 

import 'bootstrap/dist/css/bootstrap.min.css';
import "datatables.net-dt/js/dataTables.dataTables" 
import "datatables.net-dt/css/jquery.dataTables.min.css"

export interface INojavascriptWebPartProps {
  userListName: string;
  countryListName: string;
}

export default class NojavascriptWebPart extends BaseClientSideWebPart<INojavascriptWebPartProps> {

  public render(): void {
    this.domElement.innerHTML = `
              <div class="container-fluid">
              <form>
                <div class="form-group">
                  <label for="userid" class="col-3 col-form-label">User ID</label>
                  <div class="col-10">
                    <input type="text" class="form-control" id="userid" placeholder="User ID">
                  </div>
                </div>
                <div class="form-group">
                  <label for="name" class="col-3 col-form-label">Name</label>
                  <div class="col-10">
                    <input type="text" class="form-control" id="name" placeholder="Name">
                  </div>
                </div>
                <div class="form-group">
                  <label for="email" class="col-3 col-form-label">Email</label>
                  <div class="col-10">
                    <input type="email" class="form-control" id="email" placeholder="name@example.com">
                  </div>
                </div>
                <div class="form-group">
                  <label for="cno" class="col-3 col-form-label">Contact Number</label>
                  <div class="col-10">
                    <input type="text" class="form-control" id="cno" placeholder="Contact Number">
                  </div>
                </div>
                <div class="form-group">
                  <label for="country" class="col-3 col-form-label">Country</label>
                  <div class="col-10">
                    <select class="form-control" id="country">
                      <option>-- Select --</option>
                    </select>
                  </div>
                </div>
                <div class="form-group">
                  <label for="dob" class="col-3 col-form-label">Date of Birth</label>
                  <div class="col-10">
                    <input type="date" class="form-control" id="dob" placeholder="Contact Number">
                  </div>
                </div>
                <button type="button" class="btn btn-primary m-3 float-left" id="btn_submit">Submit</button>
              </form>
              <table class="table mt-3" id="tbl_users">
                <thead class="thead-dark">
                  <tr>
                    <th scope="col">User ID</th>
                    <th scope="col">Name</th>
                    <th scope="col">EMail</th>
                    <th scope="col">Contact Number</th>
                    <th scope="col">Country</th>
                    <th scope="col">Date of Birth</th>
                  </tr>
                </thead>
                <tbody>
                  
                </tbody>
              </table>
          </div>`;
          document.getElementById("btn_submit").addEventListener('click',this.AddUsers.bind(this));
          this.GetCountries();
          this.GetUsers();
  }

  private GetCountries() {

    let url = this.context.pageContext.web.absoluteUrl +  "/_api/web/lists/getbytitle('"+this.properties.countryListName+"')/items?$select=Id,Title";
    this.context.spHttpClient.get(url,SPHttpClient.configurations.v1).then(res => {
      res.json().then(responseJSON => {

        console.log(responseJSON);
        let values = responseJSON.value;
        values.forEach(element => {
          $("#country").append(
            $("<option></option>").attr('value',element.Id).html(element.Title)
          );
        });
      })
    });
    
  }

  private GetUsers() {
    let url = this.context.pageContext.web.absoluteUrl +  "/_api/web/lists/getbytitle('"+this.properties.userListName+"')/items?$expand=Country&$select=*, Country/Title";
    
    $.ajax({
      url : url,
      headers : {
        Accept: 'application/json;odata=nometadata'
      },
      type : "GET",
      success : function(data){

        let values = data.value;


        ($("#tbl_users") as any).dataTable({
          data : values,
          "bDestroy": true,
          "responsive": true,
          "autoWidth": false,
          columns : [
            {
              data : 'Title'
            },
            {
              data : 'Name'
            },
            {
              data : 'Email'
            },
            {
              data : 'ContactNo'
            },
            {
              data: 'Country.Title',
            },
            {
              data : 'DateOfBirth',
              render : function(data,row,full){
                let date = new Date(data).toISOString().split('T')[0];
                return date;
              }
            }
          ]
        });

        console.log(values);


      },
      error : function(error){
        console.error(error);
      }
    })

  }

  private AddUsers() {

    let metadata = {};
    metadata["Title"] = $("#userid").val();
    metadata["Name"] = $("#name").val();
    metadata["Email"] = $("#email").val();
    metadata["ContactNo"] = $("#cno").val();
    metadata["CountryId"] = $("#country").val();
    metadata["DateOfBirth"] = $("#dob").val(); //yyyy-mm-dd

    let url = this.context.pageContext.web.absoluteUrl +  "/_api/web/lists/getbytitle('"+this.properties.userListName+"')/items";
    
    const header = {  
      'Accept': 'application/json;odata=nometadata',  
      'Content-type': 'application/json;odata=nometadata',  
      'odata-version': ''  
    };

    const httpClientOptions: IHttpClientOptions = {  
      body: JSON.stringify(metadata),  
      headers: header  
    };

    this.context.spHttpClient.post(url,SPHttpClient.configurations.v1,httpClientOptions).then((res) => {
      res.json().then(response => {
        alert("Item Created Successfully");
      },error => {
        console.error(error);
      })
    })
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: "Webpart List Names",
              groupFields: [
                PropertyPaneTextField('userListName', {
                  label : "User List Name",
                  value : "Users"
                }),
                PropertyPaneTextField('countryListName', {
                  label: "Country List Name",
                  value : ""
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

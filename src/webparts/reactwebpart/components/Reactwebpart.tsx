import * as React from 'react';
import styles from './Reactwebpart.module.scss';
import { IReactwebpartProps } from './IReactwebpartProps';
import { SPHttpClient, SPHttpClientResponse, SPHttpClientConfiguration,IHttpClientOptions,DigestCache,IDigestCache } from '@microsoft/sp-http';
import { escape } from '@microsoft/sp-lodash-subset';

import * as $ from "jquery"; 

import 'bootstrap/dist/css/bootstrap.min.css';
import "datatables.net-dt/js/dataTables.dataTables" 
import "datatables.net-dt/css/jquery.dataTables.min.css"
import { data } from 'jquery';

interface IReactwebpartStates{
  countries : any[];
  users : any[];
}

export default class Reactwebpart extends React.Component<IReactwebpartProps, IReactwebpartStates> {
  constructor(props){
    super(props);
    this.state = {
      countries : [],
      users : []
    }
  }

  componentDidMount(){
    // this.GetCountries();
    // this.GetUsers();
    this.GetStudent();
  }

  private GetStudent = async () => {

    console.log("Started"); //1

    const respose = await $.ajax({
      url : "https://qantler.sharepoint.com/sites/practice/LeaveManagementSystem/_api/web/lists/getbytitle('StudentList')/items?$select=Title,RollNo,DeptId,LeaveTypeId,FromDate,ToDate,TotalNoOfDays,Id,UsersId,Dept/Title,LeaveType/Title,Users/Name,Users/EMail,Users/Title&$expand=Dept,LeaveType,Users",
      type : 'GET',
      headers : {
        'accept' : 'application/json;odata=verbose'
      }
      // success : (data) => {
        //respon = data;
      //   console.log(data); // 3
      // },
      // error : (error) => {
      //   console.error(error); //3
      // }
    });

    console.log(respose); //2

    console.log("GET Method Ended"); //3

    this.addUsers();

  }

  private async addUsers(){

    const digestCache: IDigestCache = this.props.context.serviceScope.consume(DigestCache.serviceKey);
    const digest: string = await digestCache.fetchDigest(this.props.context.pageContext.web.serverRelativeUrl);
    let metadata = {
      "__metadata" : {"type" : "SP.Data.StudentListListItem"},
      "Title" : "Student Name",
      "RollNo" : "004",
      "DeptId" : 3,
      "LeaveTypeId": 5,
      "UsersId": { 'results': [143] }
      //"UsersId" : [143]
    }
    const response = await $.ajax({
      url : "https://qantler.sharepoint.com/sites/practice/LeaveManagementSystem/_api/web/lists/getbytitle('StudentList')/items",
      type: "POST",
      headers: {
        "Accept": "application/json;odata=verbose",
        //Content-Type header: Specifies the format of the data that the client is sending to the server
        "Content-Type": "application/json;odata=verbose",
        "X-RequestDigest": digest
      },
      data : JSON.stringify(metadata)
    });

    console.log(response);
    console.log("POST method ended");
  }

  private GetCountries() {

    let url = this.props.context.pageContext.web.absoluteUrl +  "/_api/web/lists/getbytitle('"+this.props.countryListName+"')/items?$select=Id,Title";
    this.props.context.spHttpClient.get(url,SPHttpClient.configurations.v1).then(res => {
      res.json().then(responseJSON => {

        console.log(responseJSON);
        let values = responseJSON.value;
        this.setState({ countries : values });
        // values.forEach(element => {
        //   $("#country").append(
        //     $("<option></option>").attr('value',element.Id).html(element.Title)
        //   );
        // });
      })
    });
    
  }

  private GetUsers() {

    let _this = this;

    let url = this.props.context.pageContext.web.absoluteUrl +  "/_api/web/lists/getbytitle('"+this.props.userListName+"')/items?$expand=Country&$select=*, Country/Title";
    
    $.ajax({
      url : url,
      headers : {
        Accept: 'application/json;odata=nometadata'
      },
      type : "GET",
      success : function(data){

        let values = data.value;

        _this.setState({users : values});

        ($("#tbl_users") as any).dataTable({
          
          // "bDestroy": true,
          // "responsive": true,
          // "autoWidth": false
        });



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

    let url = this.props.context.pageContext.web.absoluteUrl +  "/_api/web/lists/getbytitle('"+this.props.userListName+"')/items";
    
    const header = {  
      'Accept': 'application/json;odata=nometadata',  
      'Content-type': 'application/json;odata=nometadata',  
      'odata-version': ''  
    };

    const httpClientOptions: IHttpClientOptions = {  
      body: JSON.stringify(metadata),  
      headers: header  
    };

    this.props.context.spHttpClient.post(url,SPHttpClient.configurations.v1,httpClientOptions).then((res) => {
      res.json().then(response => {
        alert("Item Created Successfully");
      },error => {
        console.error(error);
      })
    })
  }

  public render(): React.ReactElement<IReactwebpartProps> {
    return (
      <div className="container-fluid">
          <form>
            <div className="form-group">
              <label htmlFor="userid" className="col-3 col-form-label">User ID</label>
              <div className="col-10">
                <input type="text" className="form-control" id="userid" placeholder="User ID" />
              </div>
            </div>
            <div className="form-group">
              <label htmlFor="name" className="col-3 col-form-label">Name</label>
              <div className="col-10">
                <input type="text" className="form-control" id="name" placeholder="Name" />
              </div>
            </div>
            <div className="form-group">
              <label htmlFor="email" className="col-3 col-form-label">Email</label>
              <div className="col-10">
                <input type="email" className="form-control" id="email" placeholder="name@example.com" />
              </div>
            </div>
            <div className="form-group">
              <label htmlFor="cno" className="col-3 col-form-label">Contact Number</label>
              <div className="col-10">
                <input type="text" className="form-control" id="cno" placeholder="Contact Number" />
              </div>
            </div>
            <div className="form-group">
              <label htmlFor="country" className="col-3 col-form-label">Country</label>
              <div className="col-10">
                <select className="form-control" id="country">
                  <option>-- Select --</option>
                  {/* {
                    this.state.countries.map((country) => {
                      return <option value={country.Id}>{country.Title}</option> //<option value="1">India</option>
                    })
                  } */}
                </select>
              </div>
            </div>
            <div className="form-group">
              <label htmlFor="dob" className="col-3 col-form-label">Date of Birth</label>
              <div className="col-10">
                <input type="date" className="form-control" id="dob" placeholder="Contact Number" />
              </div>
            </div>
            <button type="button" className="btn btn-primary m-3 float-left" id="btn_submit" onClick={this.AddUsers.bind(this)}>Submit</button>
          </form>
          <table className="table mt-3" id="tbl_users">
            <thead className="thead-dark">
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

              {/* {
                this.state.users.map((user) => {
                 // <tr><td>001</td>Ganesh</td>....</tr>
                  return (
                    <tr>
                      <td>{user.Title}</td>
                      <td>{user.Name}</td>
                      <td>{user.Email}</td>
                      <td>{user.ContactNo}</td>
                      <td>{user.Country.Title}</td>
                      <td>{ new Date(user.DateOfBirth).toISOString().split('T')[0] }</td>

                    </tr>
                  )

                })
              } */}
            </tbody>
          </table>
        </div>
    );
  }
}

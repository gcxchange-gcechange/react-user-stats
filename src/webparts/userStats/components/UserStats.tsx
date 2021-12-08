import * as React from 'react';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import {
  autobind,
  DetailsList,
  Stack,
  DefaultButton
} from 'office-ui-fabric-react';

import styles from './UserStats.module.scss';
import { IUserStatsProps } from './IUserStatsProps';
import { IUserStatsState } from './IUserStatsState';
import { escape } from '@microsoft/sp-lodash-subset';

export default class UserStats extends React.Component<IUserStatsProps, IUserStatsState> {

constructor(props: IUserStatsProps, state: IUserStatsState) {
  super(props);

  this.state = {
    allUsers: [],
    countAllUsers: [],
    countAllUsersTest: [],
    groupsDelta: [],
  }
}

@autobind
private getUserDetails(): void {
  // msgraph call for users
  // Get createdDateTime of users in AD
  this.props.context.msGraphClientFactory  
    .getClient()  
    .then((client: MSGraphClient): void => {  
      // Get user information from the Microsoft Graph  
      //createdDateTime

      client  
        .api('/users')
        .version("v1.0")
        .select(["id","displayName","mail","createdDateTime"]) 
        .get((error, result: any, rawResponse?: any) => {  
          // handle the response  
          if (error) {  
            console.log(error);
            return;  
          }  
          // console.log(result);
          var todayDate = Date.now();
          var allDates = []
          // get all of the account creation dates and format them into new array
          // ex: 2021-10 (yyyy-mm)
          // then count the duplicates to get # of registrations per month 
          var dates = result.value.map(date => {
            var splitDate = date.createdDateTime.split("-");
            allDates.push(`${splitDate[0]}-${splitDate[1]}`)
            return date.createdDateTime
          });
          console.log(todayDate);
          console.log(allDates);
          var duplicateCount = {};
          allDates.forEach(e => duplicateCount[e] = duplicateCount[e] ? duplicateCount[e] + 1 : 1);
          var resultTest = Object.keys(duplicateCount).map(e => {return {key:e, count:duplicateCount[e]}});
          console.log(resultTest);

          this.setState({
            allUsers: dates,
            countAllUsers: resultTest,
            countAllUsersTest: allDates,
          })
          console.log(this.state.countAllUsersTest);
        });  
    });  
}

@autobind
private getGroupDetails(): void {  
  // msgraph Get group member "delta"
  // Try also '/groups/delta'
  // Maybe: GET https://graph.microsoft.com/v1.0/groups/delta?$select=displayName,description,members
  // https://docs.microsoft.com/en-us/graph/delta-query-groups
  // Get count: https://graph.microsoft.com/v1.0/groups/02bd9fd6-8f93-4758-87c3-1fb73740a315/members?$count=true
  // delta may not work how we intended
  this.props.context.msGraphClientFactory  
    .getClient()  
    .then((client: MSGraphClient): void => {  
      // Get user information from the Microsoft Graph  
      //createdDateTime

      client  
        .api('/groups/delta')
        .version("v1.0")
        .select(["id","displayName","members"]) 
        .get((error, result: any, rawResponse?: any) => {  
          // handle the response  
          if (error) {  
            console.log(error);
            return;  
          }
          console.log(result);
          var formatRes = [];
          result.value.map(e => {
            //Loop and call and pray it doesn't crash!
            var itemCount = this.getGroupCount(e.id);
            // This kind of worked, but only after the push, perhaps I need to wait
            formatRes.push({
              key: e.id,
              displayName: e.displayName,
              // I can't grab members@delta without an error
              // count: e.members@delta.length,
              count: itemCount,
            })
          });
          console.log(formatRes);
          this.setState({
            groupsDelta: result.value,
          })
        });  
    });  
}

@autobind
private getGroupCount(groupID): void {  
  // msgraph Get a group member count
  this.props.context.msGraphClientFactory  
    .getClient()  
    .then((client: MSGraphClient): void => {  
      // Get user information from the Microsoft Graph  
      //createdDateTime

      client  
        .api(`/groups/${groupID}/members`)
        .version("v1.0") 
        .get((error, result: any, rawResponse?: any) => {  
          // handle the response  
          if (error) {  
            console.log(error);
            return;  
          }  
          // console.log(result);
          // I could just loop through this and get count and build a new array?
          console.log(result.value.length);
          return(result.value.length);
        });  
    });  
}

@autobind
private getAadUsers(): void {
  this.props.context.aadHttpClientFactory
    .getClient('')
    .then((client: AadHttpClient): void => {
      client
        .get('', AadHttpClient.configurations.v1)
        .then((response: HttpClientResponse): Promise<any> => {
          console.log(response.json());
          return response.json();
        })
    });
}

@autobind
private async getAsyncUsers() {
  // get async users
  const response = await fetch('', {
    method: 'POST',
    headers: {
        'Content-Type':'application/json'
    }}).then((data) => {
      return data
    })
    .catch((error) => {
      console.log(error);
      return
    })
  return await response;
}

@autobind
private async getAsyncGroups() {
  // get async groups
  const response = await fetch('', {
    method: 'POST',
    headers: {
        'Content-Type':'application/json'
    }}).then((data) => {
      return data
    })
    .catch((error) => {
      console.log(error);
      return
    })
  return await response;
}


componentDidMount() {
  // Call APIs on mount
  // this.getUserDetails();
  // this.getGroupDetails()
  this.getAadUsers();
  /*
  this.getAsyncGroups().then(u => {
      console.log('Async Fetch:');
      console.log(u);
  })
  */
}




  public render(): React.ReactElement<IUserStatsProps> {
    var testItem = [
      {key: "Loading...", count: 10},
    ]
    var testCols = [
      { key: 'key', name: 'Year-Month', fieldName: 'key', minWidth: 85, maxWidth: 90, isResizable: true },
      { key: 'column2', name: 'New Users', fieldName: 'count', minWidth: 200, maxWidth: 225, isResizable: true },
    ]
    var testDepart = [
      {key: "TBS", value:100},
      {key: "SSC", value:75},
      {key: "TEST", value:45}
    ]
    var departCols = [
      { key: 'key', name: 'ID', fieldName: 'id', minWidth: 20, maxWidth: 20, isResizable: true },
      { key: 'column2', name: 'Department', fieldName: 'displayName', minWidth: 200, maxWidth: 225, isResizable: true },
      { key: 'column3', name: 'Member Count', fieldName: 'count', minWidth: 100, maxWidth: 125, isResizable: true },
      
    ]
    return (
      <div className={ styles.userStats }>
        <div>
          <div>
            <div>
              <h1>User Stats</h1>
              <Stack horizontal disableShrink>
                <div className={ styles.statsHolder }>
                  <h2>Total Users:</h2>
                  <div className={ styles.userCount }>{this.state.allUsers.length}</div>  
                </div>
                <div>
                  <h2>Breakdown by Month</h2>
                  <div>
                    <DetailsList
                      items={this.state.countAllUsers ?  this.state.countAllUsers : testItem}
                      compact={true}
                      columns={testCols}
                    />
                  </div>
                </div>
              </Stack>
              <h2>Groups / Departments</h2>
              <Stack>
                <div>
                  <DetailsList
                    items={this.state.groupsDelta ? this.state.groupsDelta : testDepart}
                    compact={true}
                    columns={departCols}
                  />
                </div>
              </Stack>
              <h2>Testing</h2>
              <Stack>
                <ul>
                  {
                    this.state.groupsDelta ?
                    this.state.groupsDelta.map(group => (
                      <li key={group.id}>
                        <DefaultButton
                          text={group.displayName}
                          onClick={() => {
                            this.getGroupCount(group.id);
                          }}
                        />
                      </li>
                    )) : <li>NOTHING</li>
                  }
                </ul>
              </Stack>
            </div>
          </div>
        </div>
      </div>
    );
  }
}

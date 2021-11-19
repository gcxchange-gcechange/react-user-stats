import * as React from 'react';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import {
  autobind,
  DefaultButton
} from 'office-ui-fabric-react';
// import styles from './UserStats.module.scss';
import { IUserStatsProps } from './IUserStatsProps';
import { IUserStatsState } from './IUserStatsState';
import { escape } from '@microsoft/sp-lodash-subset';

export default class UserStats extends React.Component<IUserStatsProps, IUserStatsState> {

constructor(props: IUserStatsProps, state: IUserStatsState) {
  super(props);

  this.state = {
    allUsers: [],
    countAllUsers: [],
    groupsDelta: [],
  }
}

@autobind
private getUserDetails(): void {
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
          // var count = [];
          var count = allDates.reduce(function(prev, cur) {
            prev[cur] = (prev[cur] || 0) + 1;
            return prev;
          }, {});
          console.log(count)
          // allDates.forEach(function(i) { return count[i] = (count[i]||0) + 1;});
          this.setState({
            allUsers: dates,
            countAllUsers: count,
          })
          console.log(this.state.countAllUsers);
        });  
    });  
}

@autobind
private getGroupDetails(): void {  
  // Get group member "delta"
  // Try also '/groups/delta'
  // Maybe: GET https://graph.microsoft.com/v1.0/groups/delta?$select=displayName,description,members
  // https://docs.microsoft.com/en-us/graph/delta-query-groups
  // Get count: https://graph.microsoft.com/v1.0/groups/02bd9fd6-8f93-4758-87c3-1fb73740a315/members?$count=true
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
          this.setState({
            groupsDelta: result.value,
          })
        });  
    });  
}

@autobind
private getGroupCount(groupID): void {  
  // Get a group member count
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
          alert(result.value.length);
        });  
    });  
}

componentDidMount() {
  // Call APIs on mount
  this.getUserDetails();
  this.getGroupDetails()
}

  public render(): React.ReactElement<IUserStatsProps> {
    return (
      <div>
        <div>
          <div>
            <div>
              <h1>User Stats</h1>
              <h2>Total Users: {this.state.allUsers.length}</h2>
              <h2>Month created break down</h2>
              <div>
                <ul>
                  <li>HELLO?!</li>
                  {
                    this.state.countAllUsers ?
                    /*
                    this.state.countAllUsers.map(item =>(
                      <li>{item}</li>
                    )) 
      */
                      <li>YES TEST</li>
                    : <li>NOTHING</li>
                  
                  }
                </ul>
              </div>
              <h2>Groups / Departments</h2>
              <ul>
                {
                  this.state.groupsDelta ?
                  this.state.groupsDelta.map(group => (
                    <li key={group.id}>
                      <DefaultButton
                        text={group.displayName}
                        onClick={() => {
                          // alert(`I will fetch: ${group.id}`);
                          this.getGroupCount(group.id);
                        }}
                      />
                      Members Delta:
                    </li>
                  )) : <li>NOTHING</li>
                }
              </ul>
            </div>
          </div>
        </div>
      </div>
    );
  }
}

import * as React from 'react';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import {
  autobind,
  DetailsList,
  Stack
} from 'office-ui-fabric-react';

import styles from './UserStats.module.scss';
import { IUserStatsProps } from './IUserStatsProps';
import { IUserStatsState } from './IUserStatsState';

export default class UserStats extends React.Component<IUserStatsProps, IUserStatsState> {

constructor(props: IUserStatsProps, state: IUserStatsState) {
  super(props);

  this.state = {
    allUsers: [],
    countAllUsers: [],
    groupsDelta: [],
    communityCount: [],
    testArray: [],
    userLoading: true,
    groupLoading: true
  }
}

// User Stats Call
@autobind
private getAadUsers(): void {
  this.props.context.aadHttpClientFactory
    .getClient('')
    .then((client: AadHttpClient) => {
      client
        .get('', AadHttpClient.configurations.v1)
        .then((response: HttpClientResponse): Promise<any> => {
          response.json().then(((r) => {
            // console.log(r);
            // Format dates to Year-Month (ex 2021-10)
            var allDates = [];
            var dates = r.map(date => {
              var splitDate = date.createDateTime.split("-");
              allDates.push(`${splitDate[0]}-${splitDate[1]}`)
              return date.createDateTime
            });
            // Count duplicates 
            var duplicateCount = {};
            allDates.forEach(e => duplicateCount[e] = duplicateCount[e] ? duplicateCount[e] + 1 : 1);
            var resultTest = Object.keys(duplicateCount).map(e => {return {key:e, count:duplicateCount[e]}});
            // Sort the dates
            resultTest.sort(function (a,b) {
              var keyA = a.key.replace('-', '');
              var keyB = b.key.replace('-', '');
              return parseInt(keyB) - parseInt(keyA);
            })
            // Set the state
            this.setState({
              allUsers: dates,
              countAllUsers: resultTest,
              userLoading: false
            })
          }))
          return response.json();
        })
    });
}

// Group Stats Call
@autobind
private getAadGroups(): void {
  this.props.context.aadHttpClientFactory
    .getClient('')
    .then((client: AadHttpClient) => {
      client
        .get('', AadHttpClient.configurations.v1)
        .then((response: HttpClientResponse): Promise<any> => {
          response.json().then(((r) => {
            // console.log(r);
            // Get a count of communities (Unified group type)
            var totalCommunities = []
            r.map(c => {
              if (c.groupType[0] === 'Unified') {
                totalCommunities.push(c.displayName);
              }
            })
            // console.log(totalCommunities);
            // Filter out community groups by their type to leave mostly departments
            var filteredR = r.filter(item => item.groupType[0] !== 'Unified');
            // Set the state
            // console.log(filteredR);
            var allDepartments = [];
            filteredR.map(s => {
              var splitS = s.displayName.split("_")
              if (splitS.length > 1) {
                allDepartments.push(`${splitS[1]} - ${s.countMember}`);
              }
            });
            this.setState({
              groupsDelta: filteredR,
              communityCount: totalCommunities,
              testArray: allDepartments,
              groupLoading: false
            });
          }));
          return response.json();
        })
  })
}

componentDidMount() {
  // Call APIs on mount
  this.getAadUsers();
  this.getAadGroups();
}

  public render(): React.ReactElement<IUserStatsProps> {
    // Format detail lists columns
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
      { key: 'column2', name: 'Department', fieldName: 'displayName', minWidth: 200, maxWidth: 225, isResizable: true },
      { key: 'column3', name: 'Member Count', fieldName: 'countMember', minWidth: 100, maxWidth: 125, isResizable: true },
      
    ]
    return (
      <div className={ styles.userStats }>
        <div>
          <div>
            <div>
              <h1>User Stats</h1>
              <div>
                {this.state.userLoading && 'Loading Users...'}
              </div>
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
              <h2>Departments and Communities</h2>
              <div>
                {this.state.groupLoading && 'Loading Groups...'}
              </div>
              <Stack horizontal disableShrink>
                <div className={ styles.statsHolder }>
                  <h2>Total Communities:</h2>
                  <div className={ styles.userCount }>{this.state.communityCount.length}</div>  
                </div>
                <div>
                  <h2>Department count</h2>
                  {
                    /**
                     * <DetailsList
                        items={this.state.groupsDelta ? this.state.groupsDelta : testDepart}
                        compact={true}
                        columns={departCols}
                        />
                     */
                    this.state.testArray && 
                    this.state.testArray.map(d => {
                       return <div className={ styles.departList } key={d.key}>{d}</div>
                    })
                  }
                  
                </div>
              </Stack>
            </div>
          </div>
        </div>
      </div>
    );
  }
}

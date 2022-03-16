import * as React from 'react';
import { AadHttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import {
  autobind,
  DefaultButton,
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
      filteredDepartments: [],
      userLoading: true,
      groupLoading: true
    }
  }

  // User Stats Call
  @autobind
  private getAadUsers(): void {
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: `{ "containerName": "userstats" }`
    };

    this.props.context.aadHttpClientFactory
      .getClient("9f778828-4248-474a-aa2b-ade60459fb87")
      .then((client: AadHttpClient) => {
        client
          .post("https://appsvc-function-dev-stats-dotnet001.azurewebsites.net/api/RetreiveData", AadHttpClient.configurations.v1, postOptions)
          .then((response: HttpClientResponse): Promise<any> => {
            response.json().then(((r) => {

              var allDays = [];
              var allMonths = [];

              var dates = r.map(date => {
                var splitDate = date.creationDate.split("-");

                allMonths.push(`${splitDate[0]}-${splitDate[1]}`)
                allDays.push(`${splitDate[0]}-${splitDate[1]}-${splitDate[2].split("T")[0]}`);

                return date.creationDate
              });

              // Count duplicates 
              var duplicateMonthCount = {};
              allMonths.forEach(e => duplicateMonthCount[e] = duplicateMonthCount[e] ? duplicateMonthCount[e] + 1 : 1);
              var duplicateDayCount = {};
              allDays.forEach(e => duplicateDayCount[e] = duplicateDayCount[e] ? duplicateDayCount[e] + 1 : 1);

              var resultByMonth = Object.keys(duplicateMonthCount).map(e => {return {key:e, count:duplicateMonthCount[e], report: {
                title: "user-stats-" + e,
                csv: [
                  ["Date", "New Users"]
                ]
              }}});
              var resultByDay = Object.keys(duplicateDayCount).map(e => {return {key:e, count:duplicateDayCount[e]}});

              // Sort the dates
              resultByMonth.sort(function (a,b) {
                var keyA = a.key.replace('-', '');
                var keyB = b.key.replace('-', '');
                return parseInt(keyB) - parseInt(keyA);
              });

              resultByDay.sort(function (a,b) {
                var keyA = a.key.split('-').join('');
                var keyB = b.key.split('-').join('');
                return parseInt(keyB) - parseInt(keyA);
              });

              //console.log("By Month");
              //console.log(resultByMonth);

              //console.log("By Day");
              //console.log(resultByDay);

              // Build the csv for each month
              resultByMonth.forEach(month => {

                let index = 0;
                while(true) {
                  if(resultByDay[index] == undefined) { index--; break; }
                  if(resultByDay[index].key.indexOf(month.key) === -1) { break; }
                  
                  month.report.csv.push([resultByDay[index].key, resultByDay[index].count]);
                  index++;
                }

                // Remove the days we've already added
                resultByDay.splice(0, index);
              });

              //console.log("Months List");
              //console.log(resultByMonth);

              // Set the state
              this.setState({
                allUsers: dates,
                countAllUsers: resultByMonth,
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

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: `{ "containerName": "groupstats" }`
    };

    this.props.context.aadHttpClientFactory
      .getClient("9f778828-4248-474a-aa2b-ade60459fb87")
      .then((client: AadHttpClient) => {
        client
          .post("https://appsvc-function-dev-stats-dotnet001.azurewebsites.net/api/RetreiveData", AadHttpClient.configurations.v1, postOptions)
          .then((response: HttpClientResponse): Promise<any> => {
            response.json().then(((r) => {
              //console.log(r);
              // Get a count of communities (Unified group type)
              var totalCommunities = []
              r.map(c => {
                if (c.groupType[0] === 'Unified') {
                  totalCommunities.push(c.displayName);
                }
              })
              //console.log(totalCommunities);
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
                filteredDepartments: allDepartments,
                groupLoading: false
              });
            }));
            return response.json();
          })
    })
  }

  // https://stackoverflow.com/a/14966131
  private downloadCSV(title: string, data: any) {
    let content = "data:text/csv;charset=utf-8,";

    data.forEach(function(rowArray) {
      let row = rowArray.join(",");
      content += row + "\r\n";
    });

    var encodedUri = encodeURI(content);
    var link = document.createElement("a");

    link.setAttribute("href", encodedUri);
    link.setAttribute("download", title + ".csv");

    document.body.appendChild(link);

    link.click();
  }

  componentDidMount() {
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
      { key: 'report', name: 'Report', fieldName: 'report', minWidth: 85, maxWidth: 90, isResizable: true, onRender: (item: any) => (
        <DefaultButton
          onClick={() => {
            this.downloadCSV(item.report.title, item.report.csv);
          }}
        >
          Download
        </DefaultButton>), 
      },
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
                    this.state.filteredDepartments && 
                    this.state.filteredDepartments.map(d => {
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

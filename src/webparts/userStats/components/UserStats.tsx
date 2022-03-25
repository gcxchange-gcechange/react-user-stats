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

  // *** replace these ***
  private clientId = '';
  private url = '';
  // *********************

  constructor(props: IUserStatsProps, state: IUserStatsState) {
    super(props);

    this.state = {
      allUsers: [],
      countByMonth: [],
      groupsDelta: [],
      communityCount: [],
      communitiesPerDay: [],
      communitiesPerMonth: [],
      filteredDepartments: [],
      userLoading: true,
      groupLoading: true,
      totalactiveuser: ""
    }
  }

  // User Stats Call
  @autobind
  private async getAadUsers(): Promise<any> {

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: `{ "containerName": "userstats" }`
    };

    this.props.context.aadHttpClientFactory
      .getClient(this.clientId)
      .then((client: AadHttpClient) => {
        client
          .post(this.url, AadHttpClient.configurations.v1, postOptions)
          .then((response: HttpClientResponse): Promise<any> => {
            response.json().then(((r) => {

              var allDays = [];
              var allMonths = [];

              var allUserCount = r.map(date => {
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

              var resultByMonth = Object.keys(duplicateMonthCount).map(e => {return {key:e, count:duplicateMonthCount[e], communities: 0, report: {
                title: "gcx-stats-" + e,
                csv: [
                  ["Date", "New Users", "New Communities"]
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
                  
                  month.report.csv.push([resultByDay[index].key, resultByDay[index].count, 0]);
                  index++;
                }

                // Remove the days we've already added
                resultByDay.splice(0, index);
              });

              //console.log("Months List");
              //console.log(resultByMonth);

              // Set the state
              this.setState({
                allUsers: allUserCount,
                countByMonth: resultByMonth,
                userLoading: false
              });
            }));
            
            return response.json();
          })
      });
  }

  // Group Stats Call
  @autobind
  private async getAadGroups(): Promise<any> {

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: `{ "containerName": "groupstats" }`
    };

    this.props.context.aadHttpClientFactory
      .getClient(this.clientId)
      .then((client: AadHttpClient) => {
        client
          .post(this.url, AadHttpClient.configurations.v1, postOptions)
          .then((response: HttpClientResponse): Promise<any> => {
            response.json().then(((r) => {
              //console.log(r);

              // Get a count of communities (Unified group type)
              var totalCommunities = [];
              var allMonths = [];
              r.map(c => {
                if (c.groupType[0] === 'Unified') {

                  let splitDate = c.creationDate.split(" ")[0].split("/");

                  // Format the date to match the user/csv info (mm/dd/yyyy to yyyy-mm-dd)
                  let formattedDate = splitDate[2] + "-" 
                  + (splitDate[0].length === 1 ? "0" + splitDate[0] : splitDate[0]) + "-" 
                  + (splitDate[1].length === 1 ? "0" + splitDate[1] : splitDate[1]);

                  allMonths.push(formattedDate.substring(0, 7));

                  totalCommunities.push({title: c.displayName, creationDate: formattedDate});
                }
              });

              // Sort by creation date
              totalCommunities.sort(function (a,b) {
                var keyA = a.creationDate.split('-').join('');
                var keyB = b.creationDate.split('-').join('');
                return parseInt(keyB) - parseInt(keyA);
              });

              var communitiesPerMonth = {};
              allMonths.forEach(e => communitiesPerMonth[e] = communitiesPerMonth[e] ? communitiesPerMonth[e] + 1 : 1);

              // Count duplicates to get the communities created per day
              let communitiesPerDay = {};
              totalCommunities.forEach(community => {
                communitiesPerDay[community.creationDate] = (communitiesPerDay[community.creationDate] || 0) + 1;
              });
              communitiesPerDay = Object.keys(communitiesPerDay).map((key) => [key, communitiesPerDay[key]]);

              // Filter out community groups by their type to leave mostly departments
              var filteredR = r.filter(item => item.groupType[0] !== 'Unified');

              var allDepartments = [];
              filteredR.map(s => {
                var splitS = s.displayName.split("_")
                if (splitS.length > 1) {
                  allDepartments.push(`${splitS[1]} - ${s.countMember}`);
                }
              });

              // Set the state
              this.setState({
                groupsDelta: filteredR,
                communityCount: totalCommunities,
                communitiesPerDay: communitiesPerDay,
                communitiesPerMonth: communitiesPerMonth,
                filteredDepartments: allDepartments,
                groupLoading: false,
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

  // Active user Stats Call
  @autobind
  private getAadActive(): void {

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: `{ containerName: 'activeusers' }`
    };

    this.props.context.aadHttpClientFactory
      .getClient(this.clientId)
      .then((client: AadHttpClient) => {
        client
          .post(this.url, AadHttpClient.configurations.v1, postOptions)
          .then((response: HttpClientResponse): Promise<any> => {
            response.json().then(((r) => {
              var activeusers = ""
              r.map(c => {
                activeusers = c.countActiveusers;
              })
              this.setState({
                totalactiveuser: activeusers,
              });
            }));
            return response.json();
          })
      })
  }

  componentDidMount() {
    this.getAadUsers();
    this.getAadGroups();
    this.getAadActive();
  }

  componentDidUpdate(prevProps, prevState) {
    if ((prevState.groupLoading === true && this.state.groupLoading === false) || (prevState.userLoading === true && this.state.userLoading === false)) {
      this.buildCSV();
    }
  }

  private buildCSV() {
    var monthCount = JSON.parse(JSON.stringify(this.state.countByMonth));
    var communitiesPerDay = JSON.parse(JSON.stringify(this.state.communitiesPerDay));

    for(let i = 0; i < monthCount.length; i++) {

      monthCount[i].communities = this.state.communitiesPerMonth[monthCount[i].key] ? 
      this.state.communitiesPerMonth[monthCount[i].key] : monthCount[i].communities;

      for(let c = 0;c < communitiesPerDay.length; c++) {

        let key = communitiesPerDay[c][0].substring(0, 7);

        // Check if the year-month match
        if(monthCount[i].key == key) {

          let communityDate = communitiesPerDay[c][0].split('-').join('');

          // Start at index 1 since the first index is the table header
          for(let k = 1; k < monthCount[i].report.csv.length; k++) {

            let indexDate = monthCount[i].report.csv[k][0].split('-').join('');

            // No entry exists, create one.
            if(communityDate > indexDate) {
              monthCount[i].report.csv.splice(k, 0, [communitiesPerDay[c][0], 0, communitiesPerDay[c][1]]);
              k += 2; // increment csv by 2 so we account for the new entry,
            }
            // Entry exists, add community count.
            else if (communityDate == indexDate) {
              monthCount[i].report.csv[k][2] = communitiesPerDay[c][1];
              c++; // increment community counter
            }
          }

          // Add any dates that are earlier than the earliest date in the CSV
          let earliestDate = monthCount[i].report.csv[monthCount[i].report.csv.length - 1][0].split('-').join('');
          if(communityDate < earliestDate) {
            monthCount[i].report.csv.push([communitiesPerDay[c][0], 0 , communitiesPerDay[c][1]]);
          }
        }
      }
    }
    
    this.setState({
      countByMonth: monthCount,
    });
  }

  public render(): React.ReactElement<IUserStatsProps> {
    // Format detail lists columns
    var testItem = [
      {key: "Loading...", count: 10},
    ]
    var testCols = [
      { key: 'key', name: 'Year-Month', fieldName: 'key', minWidth: 85, maxWidth: 90, isResizable: true },
      { key: 'newUsers', name: 'New Users', fieldName: 'count', minWidth: 100, maxWidth: 115, isResizable: true },
      { key: 'newCommunities', name: 'New Communities', fieldName: 'communities', minWidth: 100, maxWidth: 115, isResizable: true },
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
                      items={this.state.countByMonth ?  this.state.countByMonth : testItem}
                      compact={true}
                      columns={testCols}
                    />
                  </div>
                </div>
                <div className={styles.statsHolder}>
                  <h2>Total active Users</h2><h3>In the last 30 days:</h3>
                  <div className={styles.userCount}>{this.state.totalactiveuser}</div>
                </div>
              </Stack>
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

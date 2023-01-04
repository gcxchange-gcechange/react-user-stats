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
import * as moment from 'moment';

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
      totalactiveuser: "",
      nmb_com_member_3: 0,
      nmb_com_member_5: 0,
      nmb_com_member_10: 0,
      nmb_com_member_20: 0,
      nmb_com_member_30: 0,
      nmb_com_member_31: 0,
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

                allMonths.push(`${splitDate[0]}-${splitDate[1]}`);
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

              // Add entries up to the current date (if no new users for those months) so there are no gaps
              let today = moment(new Date());
              let currYear = today.format('YYYY');
              let currMonth = today.format('MM');

              let startYear = parseInt(resultByMonth[resultByMonth.length - 1].key.split('-')[0]);
              let startMonth = parseInt(resultByMonth[resultByMonth.length - 1].key.split('-')[1]);

              // Get the number of months from today's date to the oldest date in the list
              let monthsDifference = parseInt(currMonth) + 1 - startMonth + 12 * (parseInt(currYear) - startYear);

              let fullResults = [];
              let earliestDate = moment(startYear + '-' + startMonth);
              for(let i = 0; i < monthsDifference; i++) {

                if(i !== 0) 
                  earliestDate.add(1, 'months');
                
                let entry = this.generateEntry(earliestDate.format('YYYY'), earliestDate.format('MM'));
                fullResults.push(entry);

                for(let c = 0; c < resultByMonth.length; c++) {
                  if(fullResults[i].key == resultByMonth[c].key) {
                    fullResults[i] = resultByMonth[c];
                    break;
                  }
                }
              }

              // Set the state
              this.setState({
                allUsers: allUserCount,
                countByMonth: fullResults.reverse(),
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

    var nmb_com_member_3 = 0;
    var nmb_com_member_5 = 0;
    var nmb_com_member_10 = 0;
    var nmb_com_member_20 = 0;
    var nmb_com_member_30 = 0;
    var nmb_com_member_31 = 0;

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

                  totalCommunities.push({ title: c.displayName, creationDate: formattedDate });

                  if (c.countMember <= 3) {
                    nmb_com_member_3++
                  } else if (c.countMember <= 5) {
                    nmb_com_member_5++
                  } else if (c.countMember <= 10) {
                    nmb_com_member_10++
                  } else if (c.countMember <= 20) {
                    nmb_com_member_20++
                  } else if (c.countMember <= 30) {
                    nmb_com_member_30++
                  } else {
                    nmb_com_member_31++
                  }
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
              var allDepartmentsB2B = [];// Only depart that have a B2B group
              var allDepartmentsFinal = []; //Final array that is use

              filteredR.map(s => {
                var splitS = s.displayName.split("_")

               if (splitS.length > 1){
                  if (splitS[2] == "B2B") {
                    allDepartmentsB2B.push(`${splitS[1]} - ${s.countMember}`); //Create an array of B2B to compare
                    allDepartmentsFinal.push(`${splitS[1]} - ${s.countMember}`);// B2B are the final group
                  }
                  if (splitS[1] == "DFO") {
                    allDepartments.push(`${splitS[1]} - ${s.countMember - 12166}`); //To be remove
                  } else {
                    allDepartments.push(`${splitS[1]} - ${s.countMember}`);
                  }
                }
              });

              allDepartments.map(s => {
                var splits = s.split("-")

                if (allDepartmentsB2B.find((user) => user.includes(splits[0])) == undefined) { // If no b2b group exist for the depart, add the regular group to the final list
                  console.log(" IN B2B" + splits[0])
                  allDepartmentsFinal.push(`${s}`);
                } 
              });


              // Set the state
              this.setState({
                groupsDelta: filteredR,
                communityCount: totalCommunities,
                communitiesPerDay: communitiesPerDay,
                communitiesPerMonth: communitiesPerMonth,
                filteredDepartments: allDepartmentsFinal,
                nmb_com_member_3: nmb_com_member_3,
                nmb_com_member_5: nmb_com_member_5,
                nmb_com_member_10: nmb_com_member_10,
                nmb_com_member_20: nmb_com_member_20,
                nmb_com_member_30: nmb_com_member_30,
                nmb_com_member_31: nmb_com_member_31,
                groupLoading: false,
              });
            }));
            return response.json();
          })
    })
  }

  private generateEntry(year, month) {
    let formattedMonth = this.formatMonth(month);
    return {
      key: year + '-' + formattedMonth,
      count: 0,
      communities: 0,
      report: {
        title: 'gcx-stats-' + year + '-' + formattedMonth,
        csv: [
          ['Date', 'New Users', 'New Communities'],
          [year + '-' + formattedMonth + '-01', '0', '0']
        ]
      }
    };
  }

  private formatMonth(month) {
    return month.toString().length === 1 ? '0' + month : month;
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

    try {
      // Loop through each year-month in the list
      for(let i = 0; i < monthCount.length; i++) {

        // Update the list with the total communities for that month
        monthCount[i].communities = this.state.communitiesPerMonth[monthCount[i].key] ? 
        this.state.communitiesPerMonth[monthCount[i].key] : monthCount[i].communities;

        // Loop through the communities per day list
        for(let c = 0;c < communitiesPerDay.length; c++) {
        
          let key = communitiesPerDay[c][0].substring(0, 7);
        
          // Check if the year-month match, add them to the CSV
          if(monthCount[i].key == key) {
          
            let communityDate = communitiesPerDay[c][0].split('-').join('');
          
            // Loops through the rows in our CSV
            // Start at index 1 since the first index is the table header
            for(let k = 1; k < monthCount[i].report.csv.length; k++) {
            
              let indexDate = monthCount[i].report.csv[k][0].split('-').join('');
            
              // No entry exists, create one.
              if(communityDate > indexDate) {
                monthCount[i].report.csv.splice(k, 0, [communitiesPerDay[c][0], 0, communitiesPerDay[c][1]]);
                break;
              }
              // Entry exists, add community count.
              else if (communityDate == indexDate) {
                monthCount[i].report.csv[k][2] = communitiesPerDay[c][1];
                break;
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
    }
    catch(e) {
      console.log("Error creating CSV");
      console.log(e);
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

    var allusercountminus = this.state.allUsers.length;
    var allusercountminus2 = allusercountminus - 12166


    return (
      <div className={ styles.userStats }>
        <div>
          <div>
            <div>
              <div>
                {this.state.userLoading && 'Loading Users...'}
              </div>
              <Stack horizontal disableShrink>
                <div className={styles.statsHolder}>
                  <h2>Total Users:</h2>
                  <div className={styles.userCount}>{allusercountminus2}</div>
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
                <div className={styles.statsHolder}>
                  <h2>Total Communities:</h2>
                  <div className={styles.userCount}>{this.state.communityCount.length}</div>
                </div>
                <div className={styles.statsHolder}>
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
                      return <div className={styles.departList} key={d.key}>{d}</div>
                    })
                  }
                </div>
                <div>
                  <h2>Number of community that have:</h2>
                  <div className={styles.userCount}>3 members or less: {this.state.nmb_com_member_3}</div>
                  <div className={styles.userCount}>More than 3 members but 5 members or less:{this.state.nmb_com_member_5}</div>
                  <div className={styles.userCount}>More than 5 members but 10 members or less:{this.state.nmb_com_member_10}</div>
                  <div className={styles.userCount}>More than 10 members but 20 members or less:{this.state.nmb_com_member_20}</div>
                  <div className={styles.userCount}>More than 20 members but  30 members or less:{this.state.nmb_com_member_30}</div>
                  <div className={styles.userCount}>More than 31 members {this.state.nmb_com_member_31}</div>
                </div>
              </Stack>
            </div>
          </div>
        </div>
      </div>
    );
  }
}

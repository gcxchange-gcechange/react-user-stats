/* eslint-disable @typescript-eslint/no-explicit-any */
/*eslint no-constant-condition: ["error", { "checkLoops": false }]*/

import * as React from 'react';
import { AadHttpClient, HttpClientResponse, IHttpClientOptions } from '@microsoft/sp-http';
import {
  DatePicker,
  DayOfWeek,
  DefaultButton,
  DetailsList,
  IButtonStyles,
  Stack,
  StackItem
} from 'office-ui-fabric-react';
import styles from './UserStats.module.scss';
import { IActiveUserCount, IDomainCount, IUser, IUserStatsProps } from './IUserStatsProps';
import { IUserStatsState } from './IUserStatsState';
import * as moment from 'moment';

export default class UserStats extends React.Component<IUserStatsProps, IUserStatsState> {
  // *** replace these ***
  private clientId = '9f778828-4248-474a-aa2b-ade60459fb87';
  private url = 'https://appsvc-function-dev-stats-dotnet001.azurewebsites.net/api/RetreiveData';
  // *********************

  private domainCount: IDomainCount[] = [];
  private domainCountActive: IDomainCount[] = [];

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
      selectedDate: new Date().toLocaleDateString('en-GB').replace(/\//g, '-'),
      nmb_member_per_comm_0: 0,
      nmb_member_per_comm_3: 0,
      nmb_member_per_comm_5: 0,
      nmb_member_per_comm_10: 0,
      nmb_member_per_comm_20: 0,
      nmb_member_per_comm_21: 0,
      apiGroupData: [],
      apiUserData: [],
      siteStorage: [],
      remainingStorage: [],
      siteStorageSelectDate: new Date()
    }
  }

  private async getSiteStorage(): Promise<any> {
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");

    const day = new Date(this.state.siteStorageSelectDate);
    const dayofWeek = day.getDay(), diff = day.getDate() - dayofWeek + (dayofWeek === 0 ? 0 : 0);
    day.setDate(diff);
    console.log("DAYOFWEEK", dayofWeek);

    const getdate =  ("0" + (day.getDate())).slice(-2);
    const getMonth = ("0" + (day.getMonth() + 1)).slice(-2);
    const getYear = day.getFullYear();

    const selectReportDate = getdate + "-" + getMonth + '-' + getYear ;
     console.log("dateSeleted", selectReportDate)

    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: `{
        "containerName": "groupsitestorage",
        "selectedDate":"${selectReportDate}"
      }`
    };

        await this.props.context.aadHttpClientFactory
        .getClient(this.clientId)
        .then(async(client: AadHttpClient) => {
          await client.post(this.url, AadHttpClient.configurations.v1, postOptions)
          .then((response: HttpClientResponse): Promise<any> => {
            return response.json().then(((r) => {
              console.log("R", r);
              this.setState({siteStorage: r});

            }));
          });
        })
  }

  public bytesToGB(bytes:any): number {
    const GB = (bytes / (1000 * 1000 * 1000))
    return Math.round(GB);
  }

  public bytesToTB(bytes:any): number {
    const TB = (bytes / (1000 * 1000 * 1000 * 1000))
    return Math.round(TB * 1000) / 1000;  // round to 3 decimal places
  }

  public numberWithCommas(x:number): string {
    return x.toString().replace(/\B(?=(\d{3})+(?!\d))/g, " ");
  }

  public renderStorageTableRows(): string | JSX.Element {
    const siteStorageData = this.state.siteStorage;

    const range0To20 = 0.20;
    const range21To40 = 0.40;
    const range41To60 = 0.60;
    const range60To80 = 0.80;
    const range81To100 = 1.0;

    const results = [0,0,0,0,0];

    let usedStorage: number = 0;

    siteStorageData.forEach((item:any )=> {
      usedStorage += parseInt(item.usedStorage);

      const percentage = (item.usedStorage / item.totalStorage) * 100;

      if( percentage > 0 && percentage <= range0To20) {
        results[0]++
      } else if ( percentage > range0To20 && percentage <= range21To40 ) {
        results[1]++
      } else if ( percentage > range21To40 && percentage <= range41To60 ) {
        results[2]++
      } else if ( percentage > range41To60 && percentage <= range41To60 ) {
        results[3]++
      } else if ( percentage > range60To80 && percentage <= range81To100 ) {
        results[4]++
      }
    })

    const storageCapacity: number = this.props.storageCapacity * (1000 * 1000 * 1000 * 1000);
    const availableStorage: number = storageCapacity - usedStorage;

    let usedStorageDisplay: string;
    let availableStorageDisplay: string;

    if (this.props.storageUnit === "TB") {
      usedStorageDisplay = this.numberWithCommas(this.bytesToTB(usedStorage)) + " " + this.props.storageUnit;
      availableStorageDisplay = this.numberWithCommas(this.bytesToTB(availableStorage)) + " " + this.props.storageUnit;
    } else {
      usedStorageDisplay = this.numberWithCommas(this.bytesToGB(usedStorage)) + " " + this.props.storageUnit;
      availableStorageDisplay = this.numberWithCommas(this.bytesToGB(availableStorage)) + " " + this.props.storageUnit;
    }

    return (
      <><tr>
          <td>0 - 20%</td>
          <td>{results[0]}</td>
        </tr>
        <tr>
          <td>21-40%</td>
          <td>{results[1]}</td>
        </tr>
        <tr>
          <td>41-60%</td>
          <td>{results[2]}</td>
        </tr>
        <tr>
          <td>61-80%</td>
          <td>{results[3]}</td>
        </tr>
        <tr>
          <td>81-100%</td>
          <td>{results[4]}</td>
        </tr>
        <tr>
          <th>Used storage (total)</th>
          <th>Available storage (total)</th>
        </tr>
        <tr>
          <td>{usedStorageDisplay}</td>
          <td>{availableStorageDisplay}</td>
        </tr>
      </>
    )
  }

  public renderFolderTableRows(): JSX.Element {

    const documentData = this.state.siteStorage;

    const results = [0,0,0,0];

    console.log(documentData)
    if (documentData) {

      documentData.forEach((folder:any) => {
        //console.log("Folder", folder);
        const driveList = folder.drivesList;
        if (driveList ) {
          driveList.forEach((item:any) => {
            const folderList = item.folderListItems.length
            //console.log("Items", item.folderListItems);

              if (folderList <= 5) {
              results[0]++
              } else if(folderList >=6 && folderList <= 20){
                  results[1]++
              }
              else if (folderList >=21 && folderList <= 30){
                  results[2]++
              } else if (folderList > 31) {
                  results[3]++
              }
          })
        }

      });
    }

    return (
      <><tr>
          <td> 5 or less </td>
          <td>{results[0]}</td>
        </tr>
        <tr>
          <td> 6 - 20 </td>
          <td>{results[1]}</td>
        </tr>
        <tr>
          <td> 21 - 30 </td>
          <td>{results[2]}</td>
        </tr>
        <tr>
          <td> 31 or more </td>
          <td>{results[3]}</td>
        </tr>
      </>
    )
  }

  public renderDomainCountTableRows(): string | JSX.Element {
      const rowData = this.domainCount.sort((a,b) => a.domain.toLowerCase().localeCompare(b.domain.toLowerCase())).map((row: IDomainCount) => {
      return (
          <tr key={row.domain}>
            <td>{row.domain ? row.domain: "n/a"}</td>
            <td>{row.count}</td>
          </tr>
      )
    });

    return (<>{rowData}</>);
  }

  public renderDomainCountActiveTableRows(): string | JSX.Element {
    const rowData = this.domainCountActive.sort((a,b) => a.domain.toLowerCase().localeCompare(b.domain.toLowerCase())).map((row: IDomainCount) => {
    return (
        <tr key={row.domain}>
          <td>{row.domain ? row.domain: "n/a"}</td>
          <td>{row.count}</td>
        </tr>
    )
  });

  return (<>{rowData}</>);
}

 // User Stats Call
  private async getAadUsers(): Promise<any> {
    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: `{
        "containerName": "userstats",
        "selectedDate":"${this.state.selectedDate}"
      }`
    };

    await this.props.context.aadHttpClientFactory
      .getClient(this.clientId)
      .then(async(client: AadHttpClient) => {
        await client.post(this.url, AadHttpClient.configurations.v1, postOptions)
          .then((response: HttpClientResponse): Promise<any> => {
            return response.json().then(((r) => {

              this.setState({ apiUserData: r});

              const allDays: string[] = [];
              const allMonths: string[] = [];

              const allUserCount = r.map((date: any) => {
                const splitDate = date.creationDate.split("-");

                allMonths.push(`${splitDate[0]}-${splitDate[1]}`);
                allDays.push(`${splitDate[0]}-${splitDate[1]}-${splitDate[2].split("T")[0]}`);


                return date.creationDate
              });

              // Count duplicates
              const duplicateMonthCount: any[] = [];
              allMonths.forEach((e: any) => {
                duplicateMonthCount[e] = duplicateMonthCount[e] ? duplicateMonthCount[e] + 1 : 1;
              });

              const duplicateDayCount: any[] = [];
              allDays.forEach((e: any) => {
                duplicateDayCount[e] = duplicateDayCount[e] ? duplicateDayCount[e] + 1 : 1
              });

              const resultByMonth = Object.keys(duplicateMonthCount).map((e:any) => {
                return {
                  key:e,
                  count:duplicateMonthCount[e],
                  communities: 0,
                  report: {
                  title: "gcx-stats-" + e,
                  csv: [
                    ["Date", "New Users", "New Communities"]
                  ]
              }}});
              const resultByDay = Object.keys(duplicateDayCount).map((e:any) => {
                return {
                  key:e,
                  count:duplicateDayCount[e]
                };
              });

              // Sort the dates
              resultByMonth.sort((a,b) =>  {
                const keyA = a.key.replace('-', '');
                const keyB = b.key.replace('-', '');
                return parseInt(keyB) - parseInt(keyA);
              });

              resultByDay.sort((a,b) => {
                const keyA = a.key.split('-').join('');
                const keyB = b.key.split('-').join('');
                return parseInt(keyB) - parseInt(keyA);
              });

              // Build the csv for each month
              resultByMonth.forEach(month => {

                let index = 0;
                while(true) {
                  if(resultByDay[index] === undefined) { index--; break; }
                  if(resultByDay[index].key.indexOf(month.key) === -1) { break; }

                  month.report.csv.push([resultByDay[index].key, resultByDay[index].count, 0]);
                  index++;
                }

                // Remove the days we've already added
                resultByDay.splice(0, index);
              });


              // Add entries up to the current date (if no new users for those months) so there are no gaps

              const selectedDate = this.state.selectedDate;
              const [day,  monthFormat, year] = selectedDate.split('-');
              console.log('day:',day);
              const currYear = `${year}`;
              const currMonth = `${monthFormat}`;

              const startYear = parseInt(resultByMonth[resultByMonth.length - 1].key.split('-')[0]); //output = 2021

              const startMonth = parseInt(resultByMonth[resultByMonth.length - 1].key.split('-')[1]);// output = 10 (October)

              // Get the number of months from selected  date to the oldest date in the list
              const monthsDifference = parseInt(currMonth) + 1 - startMonth + 12 * (parseInt(currYear) - startYear);

              const fullResults = [];
              const earliestDate = moment(startYear + '-' + startMonth);

              for(let i = 0; i < monthsDifference; i++) {

                if(i !== 0)
                  earliestDate.add(1, 'months');

                const entry = this.generateEntry(earliestDate.format('YYYY'), earliestDate.format('MM'));
                fullResults.push(entry);

                for(let c = 0; c < resultByMonth.length; c++) {
                  if(fullResults[i].key === resultByMonth[c].key) {
                    fullResults[i] = resultByMonth[c];
                    break;
                  }
                }
              }

              this.domainCount.length = 0;

              r.map((user: IUser) => {
                let domain: string = "";
               
                if ((user.mail !== undefined) && (user.mail !== null)) {
                  domain = user.mail.substring(user.mail.indexOf("@") + 1);
                } else {
                  domain = "";
                }

                const idx: number = this.domainCount.findIndex((elem) => {return  elem.domain === domain;});

                if (idx > -1) {
                  this.domainCount[idx].count += 1;
                } else {
                  const dCount: IDomainCount = { domain: domain, count: 1 };
                  this.domainCount.push(dCount);
                }
              });

              // Set the state
              this.setState({
                allUsers: allUserCount,
                countByMonth: fullResults.reverse(),
                userLoading: false
              });
            }));
          })
      });
  }

  // Group Stats Call

  private async getAadGroups(): Promise<any> {

    let nmb_com_member_3 = 0;
    let nmb_com_member_5 = 0;
    let nmb_com_member_10 = 0;
    let nmb_com_member_20 = 0;
    let nmb_com_member_30 = 0;
    let nmb_com_member_31 = 0;

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: `{

        "containerName": "groupstats",
        "selectedDate":"${this.state.selectedDate}"
      }`
    };

    await this.props.context.aadHttpClientFactory
      .getClient(this.clientId)
      .then(async(client: AadHttpClient) => {
        await client.post(this.url, AadHttpClient.configurations.v1, postOptions)
          .then((response: HttpClientResponse): Promise<any> => {
            return response.json().then(((r) => {
              console.log("GroupsRes", r);

              this.setState({apiGroupData: r});

              // Get a count of communities (Unified group type)
              const totalCommunities: any[] = [];
              const allMonths:any[] = [];
              r.map((c: any) => {
                //const unified = [];
                if (c.groupType[0] === 'Unified') {

                  const splitDate = c.creationDate.split(" ")[0].split("/");


                  // Format the date to match the user/csv info (mm/dd/yyyy to yyyy-mm-dd)
                  const formattedDate = splitDate[2] + "-"
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
              totalCommunities.sort( (a,b) =>  {
                const keyA = a.creationDate.split('-').join('');
                const keyB = b.creationDate.split('-').join('');
                return parseInt(keyB) - parseInt(keyA);
              });

              const communitiesPerMonth: string|number[] = [];
              allMonths.forEach((e) => { communitiesPerMonth[e] = communitiesPerMonth[e] ? communitiesPerMonth[e] + 1 : 1 });
              // Count duplicates to get the communities created per day
              let communitiesPerDay: any [] = [];
              totalCommunities.forEach((community:any) => {
                communitiesPerDay[community.creationDate] = (communitiesPerDay[community.creationDate] || 0) + 1;
              });
              communitiesPerDay = Object.keys(communitiesPerDay).map((key:any) => [key, communitiesPerDay[key]]);
              // Filter out community groups by their type to leave mostly departments
              const filteredR = r.filter((item:any) => item.groupType[0] !== 'Unified');

              const allDepartments: string[] = [];
              const allDepartmentsB2B: string[] = [];// Only depart that have a B2B group
              const allDepartmentsFinal:string[] = []; //Final array that is use

              filteredR.map((s:any) => {
                const splitS = s.displayName.split("_")


               if (splitS.length > 1){
                  if (splitS[2] === "B2B") {
                    allDepartmentsB2B.push(`${splitS[1]} - ${s.countMember}`); //Create an array of B2B to compare
                    allDepartmentsFinal.push(`${splitS[1]} - ${s.countMember}`);// B2B are the final group
                  } else {
                    allDepartments.push(`${splitS[1]} - ${s.countMember}`);
                  }
                }
              });

              allDepartments.map(s => {
                const splits = s.split("-")

                if (allDepartmentsB2B.find((user) => user.includes(splits[0])) === undefined) { // If no b2b group exist for the depart, add the regular group to the final list
                  // console.log(" IN B2B" + splits[0]);
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
          })
        });

  }


  public getUserperCommunity( ): JSX.Element {

    const unifiedGroups = this.state.apiGroupData.filter((group:any) => group.groupType[0] === 'Unified');

    const allUsers = unifiedGroups.flatMap((item:any) => item.userlist).flat();

    const countMap = new Map();
      allUsers.forEach((value:any) =>  {
      if (countMap.has(value)) {
      countMap.set(value, countMap.get(value) + 1);
      } else {
      countMap.set(value, 1);
      }
    });

    const result = [0, 0, 0, 0, 0];

    countMap.forEach((key) => {

      if (key <= 3 ) {
        result[0]++
      }
      else if ( key <= 5) {
        result[1]++
      }
      else if ( key <= 10) {
        result[2]++
      }
      else if( key <= 20 ) {
        result[3]++
      }
      else if( key >= 21 ) {
        result[4]++
      }

    })

    return (
      <><tr>
          <td>{result[0]}</td>
          <td> None </td>
        </tr>
        <tr>
          <td>{result[1]}</td>
          <td> 1 - 3 </td>
        </tr>
        <tr>
          <td>{result[2]}</td>
          <td> 4 - 5 </td>
        </tr>
        <tr>
          <td>{result[3]}</td>
          <td> 6 - 10</td>
        </tr>
        <tr>
          <td>{result[4]}</td>
          <td> 11 - 20 </td>
        </tr>
      </>
    )


  }

  private generateEntry(year:any, month:any): any {
    const formattedMonth = this.formatMonth(month);
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

  private formatMonth(month: any) : string {
    return month.toString().length === 1 ? '0' + month : month;
  }

  // https://stackoverflow.com/a/14966131
  private downloadCSV(title: string, data: any): void {

    console.log("data", data)
    let content = "data:text/csv;charset=utf-8,";

    data.forEach((rowArray:any) => {
      const row = rowArray.join(",");
      content += row + "\r\n";
    });

    const encodedUri = encodeURI(content);
    const link = document.createElement("a");

    link.setAttribute("href", encodedUri);
    link.setAttribute("download", title + ".csv");

    document.body.appendChild(link);

    link.click();
  }

  // Active user Stats Call

  private async getAadActive(): Promise<any> {

    const requestHeaders: Headers = new Headers();
    requestHeaders.append("Content-type", "application/json");
    requestHeaders.append("Cache-Control", "no-cache");
    const postOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: `{
        containerName: 'activeusers',
        "selectedDate":"${this.state.selectedDate}"
      }`
    };

    await this.props.context.aadHttpClientFactory
      .getClient(this.clientId)
      .then(async(client: AadHttpClient) => {
        await client.post(this.url, AadHttpClient.configurations.v1, postOptions)
          .then((response: HttpClientResponse): Promise<any> => {
            return response.json().then(((r) => {
              let activeusers: string = "";

              r.map((c: IActiveUserCount )=> {
                activeusers = c.countActiveusers.toString();

                this.domainCountActive.length = 0;

                if ((c.countByDomain !== undefined) && (c.countByDomain !== null)) {
                  this.domainCountActive = c.countByDomain;
                }
              })

              this.setState({
                totalactiveuser: activeusers
              });
            }));
          })
      })
  }

  public async componentDidMount():Promise<void> {
    await this.getAadUsers();
    await this.getAadGroups();
    await  this.getAadActive();
    await this.getSiteStorage();
  }

  public async componentDidUpdate(prevProps: any, prevState: any): Promise<void>{
    if ((prevState.groupLoading === true && this.state.groupLoading === false) || (prevState.userLoading === true && this.state.userLoading === false)) {
      this.buildCSV();
    }

    if (this.state.selectedDate !== prevState.selectedDate ) {
      this.setState({
        allUsers: [],
        countByMonth: [],
        communityCount: [],
        communitiesPerDay: [],
        communitiesPerMonth: [],
        siteStorage: [],
      })

      await this.getAadUsers();
      await this.getAadGroups();
      await this.getSiteStorage();
      await this.getAadActive();
    }
  }

  private buildCSV(): void {
    const monthCount = JSON.parse(JSON.stringify(this.state.countByMonth));
    const communitiesPerDay = JSON.parse(JSON.stringify(this.state.communitiesPerDay));


    try {
      // Loop through each year-month in the list
      for(let i = 0; i < monthCount.length; i++) {
        // Update the list with the total communities for that month
        monthCount[i].communities = this.state.communitiesPerMonth[monthCount[i].key] ?
        this.state.communitiesPerMonth[monthCount[i].key] : monthCount[i].communities;

        // Loop through the communities per day list
        for(let c = 0;c < communitiesPerDay.length; c++) {

          const key = communitiesPerDay[c][0].substring(0, 7);

          // Check if the year-month match, add them to the CSV
          if(monthCount[i].key === key) {

            const communityDate = communitiesPerDay[c][0].split('-').join('');

            // Loops through the rows in our CSV
            // Start at index 1 since the first index is the table header
            for(let k = 1; k < monthCount[i].report.csv.length; k++) {

              const indexDate = monthCount[i].report.csv[k][0].split('-').join('');

              // No entry exists, create one.
              if(communityDate > indexDate) {
                monthCount[i].report.csv.splice(k, 0, [communitiesPerDay[c][0], 0, communitiesPerDay[c][1]]);
                break;
              }
              // Entry exists, add community count.
              else if (communityDate === indexDate) {
                monthCount[i].report.csv[k][2] = communitiesPerDay[c][1];
                break;
              }
            }

            // Add any dates that are earlier than the earliest date in the CSV
            const earliestDate = monthCount[i].report.csv[monthCount[i].report.csv.length - 1][0].split('-').join('');

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

  private onSelectDate = (date: Date): void => {
    //const dayofWeek = date.getDay();
    const day = ("0" + (date.getDate())).slice(-2)
    const month =  ("0" + (date.getMonth() + 1)).slice(-2);
    const year = date.getFullYear();
    const formattedSelectedDate = day + '-' + month + '-' +  year;
    // const formattedSiteStorageDate = dayofWeek + '-' + day + '-' + month + '-' +  year;


    this.setState({
      selectedDate: formattedSelectedDate,
      userLoading: true,
      groupLoading: true,
      siteStorageSelectDate: date
    });

  }

  private downloadDataFile = (dataType: string): void => {

    let data: any, fileName: string;
    const selectedDate = this.state.selectedDate;

    if (dataType === 'user') {
      data = this.state.apiUserData;
      fileName = selectedDate + "-UserStats.txt";
    } else if (dataType === 'group') {
      data = this.state.apiGroupData;
      fileName = selectedDate  + "-GroupStats.txt";
    } else if (dataType === 'siteStorage') {
      data = this.state.siteStorage;
      fileName = selectedDate  + "-SiteStorage.txt";
    }

    const dataStr =
      'data:text/json;chatset=utf-8,' +
      encodeURIComponent(JSON.stringify(data, null, 2));

    const link = document.createElement("a");
    link.setAttribute("href", dataStr);
    link.setAttribute("download", fileName)

    document.body.appendChild(link);

    link.click();

  }


  public render(): React.ReactElement<IUserStatsProps> {


    // Format detail lists columns
    const testItem = [
      {key: "Loading...", count: 10},
    ];
    const testCols = [
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
    ];
    // const testDepart = [
    //   {key: "TBS", value:100},
    //   {key: "SSC", value:75},
    //   {key: "TEST", value:45}
    // ]
    // const departCols = [
    //   { key: 'column2', name: 'Department', fieldName: 'displayName', minWidth: 200, maxWidth: 225, isResizable: true },
    //   { key: 'column3', name: 'Member Count', fieldName: 'countMember', minWidth: 100, maxWidth: 125, isResizable: true },
    // ];

    const allusercountminus = this.state.allUsers.length;

    // const verticalGapStackTokens: IStackTokens = {
    //   childrenGap: 10,
    // };

    const selectedDate = this.state.selectedDate;
    const [day, month, year] = selectedDate.split('-');
    // The converted date is now in mm-dd-yyyy format
    const convertedDate = `${month}-${day}-${year}`;

    const IconStyle: Partial<IButtonStyles> = {
      icon: {color: 'white'},
      iconHovered: { color: '#c19c00'},
      rootHovered: { color: '#c19c00'}
    }


    return (
      <div className={ styles.userStats }>
        <div>
          <div>
            <div>
              <div>
                <DatePicker
                  className = {styles.calendarFieldStyles}
                  placeholder="Select a date..."
                  ariaLabel="Select a date"
                  minDate={new Date(2000,12,30)}
                  maxDate={new Date()}
                  onSelectDate={this.onSelectDate}
                  showGoToToday= {true}
                  firstDayOfWeek={DayOfWeek.Sunday}
                  value={new Date(convertedDate)}
                />
              </div>
              <div>
                {this.state.userLoading && 'Loading Users...'}
              </div>
              <Stack horizontal disableShrink horizontalAlign="space-evenly">
                <div className={styles.statsHolder}>
                  <h2>Total Users:</h2>
                  <div className={styles.userCount}>{allusercountminus}</div>
                </div>
                <div>
                  <h2 style={{textAlign:'center'}}>Breakdown by Month</h2>
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
              <div style={{marginBottom: '30px'}}>
              <Stack horizontal disableShrink horizontalAlign="space-evenly" >
                <div className={styles.statsHolder}>
                  <h2>Total Communities:</h2>
                  <div className={styles.userCount}>{this.state.communityCount.length}</div>
                </div>
                <div className={styles.statsHolder} style={{width: '400px'}}>
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
                    this.state.filteredDepartments.map(( d:any ) => {
                      return <div className={styles.departList} key={d.key}>{d}</div>
                    })
                  }
                </div>
                <Stack >
                <div style={{overflowX: 'auto'}}>
                  <h2>Community membership count</h2>

                  <table>
                    <tr>
                      <th>Number of Community Members</th>
                      <th>Number of Communities</th>
                    </tr>
                    <tr>
                      <td>3 or less</td>
                      <td>{this.state.nmb_com_member_3}</td>
                    </tr>
                    <tr>
                      <td> 4 to 5</td>
                      <td>{this.state.nmb_com_member_5} </td>
                    </tr>
                    <tr>
                      <td>6 to 10</td>
                      <td>{this.state.nmb_com_member_10} </td>
                    </tr>
                    <tr>
                      <td>11 to 20</td>
                      <td>{this.state.nmb_com_member_20} </td>
                    </tr>
                    <tr>
                      <td>21 to 30</td>
                      <td>{this.state.nmb_com_member_30} </td>
                    </tr>
                    <tr>
                      <td>31 or more</td>
                      <td>{this.state.nmb_com_member_31} </td>
                    </tr>
                  </table>
                </div>
                <div style={{overflowX: 'auto'}}>
                  <h2>Members per Community</h2>
                  <table>
                    <thead>
                      <tr>
                        <th>Number of Members</th>
                        <th>Communities Joined</th>
                      </tr>
                    </thead>
                    <tbody>
                      {this.getUserperCommunity()}
                    </tbody>
                  </table>
                </div>

                </Stack>
              </Stack>
              </div>

              <div >
                <Stack horizontal horizontalAlign="space-evenly" verticalAlign="start" >
                  <div style={{marginBottom: "12px"}}>
                    <h2>Communities Storage Capacity</h2>
                    <table>
                      <thead>
                        <tr>
                          <th>Storage percentage Range</th>
                          <th>Number of Communities</th>
                        </tr>
                      </thead>
                      <tbody>{this.renderStorageTableRows()}</tbody>
                    </table>
                  </div>

                  <div style={{marginBottom: "12px"}}>
                    <h2>File Count per Community</h2>
                    <table>
                      <thead>
                        <tr>
                          <th>Number of Communities</th>
                          <th>Document Count</th>
                        </tr>
                      </thead>
                      <tbody>{this.renderFolderTableRows()}</tbody>
                    </table>
                  </div>
                </Stack>

                <Stack horizontal horizontalAlign="space-evenly" verticalAlign="start">
                  <div style={{marginBottom: "12px"}}>
                    <h2>Active User Count by Domain</h2>
                    <table>
                      <thead>
                        <tr>
                          <th>Domain</th>
                          <th>User Count</th>
                        </tr>
                      </thead>
                      <tbody>{this.renderDomainCountActiveTableRows()}</tbody>
                    </table>
                  </div>

                  <div style={{marginBottom: "12px"}}>
                    <h2>Total User Count by Domain</h2>
                    <table>
                      <thead>
                        <tr>
                          <th>Domain</th>
                          <th>User Count</th>
                        </tr>
                      </thead>
                      <tbody>{this.renderDomainCountTableRows()}</tbody>
                    </table>
                  </div>
                </Stack>
              </div>
              <br />
              <div className={styles.sourceFileCard}>
                <h2 style={{textAlign:'center'}}>Source Files</h2>
                <Stack horizontal horizontalAlign="space-evenly" verticalAlign="center" >
                  <StackItem align='center' >
                    <DefaultButton id="UserData" styles={IconStyle} className={styles.downloadData} iconProps={{ iconName: 'CloudDownload' }} onClick={() => this.downloadDataFile('user')}>Download User Data</DefaultButton>
                  </StackItem>
                  <StackItem align='center' >
                    <DefaultButton id="GroupData" styles={IconStyle} className={styles.downloadData} iconProps={{ iconName: 'CloudDownload' }} onClick={() => this.downloadDataFile('group')}>Download Group Data</DefaultButton>
                  </StackItem>
                  <StackItem align='center' >
                    <DefaultButton id="siteStorage" styles={IconStyle} className={styles.downloadData} iconProps={{ iconName: 'CloudDownload' }} onClick={() => this.downloadDataFile('siteStorage')}>Download Site Storage Data</DefaultButton>
                  </StackItem>
                </Stack>
              </div>
              <p>Modified Date: 05/01/2024</p>
            </div>
          </div>
        </div>
      </div>
    );
  }
}

import { DefaultButton, DetailsList } from 'office-ui-fabric-react';
import * as React from 'react';
import { IUserStatsState } from './IUserStatsState';


export interface ICSVProps {
  items: any;
}


export default class CSV extends React.Component<ICSVProps> {



 public render(): React.ReactElement<ICSVProps>{

  const {items} = this.props;


  const data = Object.keys(items).map(e => { return { date:e, count:items[e] };});

  // Sort the dates
  data.sort((a,b) =>  {
    const keyA = a.date.replace('-', '');
    const keyB = b.date.replace('-', '');
    return parseInt(keyA) - parseInt(keyB);
  });

  console.log("AI", data);
  console.log("I",items);

  const columns = [
    { key: 'key', name: 'Year-Month', fieldName:'date', minWidth: 85, maxWidth: 90, isResizable: true },
    { key: 'newUsers', name: 'New Users', fieldName: 'count', minWidth: 100, maxWidth: 115, isResizable: true },
    { key: 'newCommunities', name: 'New Communities', fieldName: 'communities', minWidth: 100, maxWidth: 115, isResizable: true },
    { key: 'report', name: 'Report', fieldName: 'report', minWidth: 85, maxWidth: 90, isResizable: true, onRender: (item: any) => (
      <DefaultButton
        onClick={() => {
          // this.downloadCSV(item.report.title, item.report.csv);
        }}
      >
        Download
      </DefaultButton>),
    },
  ];

  return (
    <>
    <div> csv component</div>
    <DetailsList
      items={data}
      compact={true}
      columns={columns}
    />
    </>


  )

  }

}

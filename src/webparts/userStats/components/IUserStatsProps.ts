import { WebPartContext } from '@microsoft/sp-webpart-base'; 

export interface IUserStatsProps {
  description: string;
  storageCapacity: number;
  storageUnit: string;
  context: WebPartContext;
}

export interface IDomainCount {
  domain: string;
  count: number;
}

export interface IUser {
  Id: string;
  creationDate: Date;
  mail: string;
}

export interface IActiveUserCount {
  name: string;
  countActiveusers: number;
  countByDomain: IDomainCount[];
}
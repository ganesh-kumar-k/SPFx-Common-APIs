import { IWebPartContext } from "@microsoft/sp-webpart-base";

export interface IReactwebpartProps {
  userListName: string;
  countryListName: string;
  context: IWebPartContext;
}

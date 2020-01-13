import { IOppInfo } from "./IOppInfo";
import { IUserInfo } from "./IUserInfo";

export interface IReadySpfxTeams1State {
  // used to show the Spinner while loading data
  loading: boolean;
  // current username, taken from the service response
  username?: string;
  // message
  Message: string;
  //Opp info
  OppInfo: IOppInfo;
  //User Info
  UserInfo: IUserInfo;

}

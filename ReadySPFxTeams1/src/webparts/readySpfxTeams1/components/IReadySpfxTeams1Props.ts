import { WebPartContext } from "@microsoft/sp-webpart-base";
import * as microsoftTeams from '@microsoft/teams-js';


export interface IReadySpfxTeams1Props {
  functionUri: string;
  needsConfiguration: boolean;
  context: WebPartContext;
  _teamContext: microsoftTeams.Context;
  configureHandler: () => void;
  errorHandler: (errorMessage: string) => void;
}

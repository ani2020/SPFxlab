import * as React from 'react';
import styles from './ReadySpfxTeams1.module.scss';
import { IReadySpfxTeams1Props } from './IReadySpfxTeams1Props';
import { IReadySpfxTeams1State } from './IReadySpfxTeams1State';
import { IOppInfo } from './IOppInfo';
import { IUserInfo } from './IUserInfo';
import { IUserInfoResponse } from './IUserInfoResponse';

import { escape } from '@microsoft/sp-lodash-subset';
import * as strings from 'ReadySpfxTeams1WebPartStrings';

import {
  autobind,
  Spinner,
  SpinnerSize,
  PrimaryButton,
  TextField,
  Label,
  DetailsList,
  DetailsListLayoutMode,
  IColumn
} from 'office-ui-fabric-react';
import { AadHttpClient, HttpClientResponse, IHttpClientOptions } from "@microsoft/sp-http";


export default class ReadySpfxTeams1 extends React.Component<IReadySpfxTeams1Props, IReadySpfxTeams1State> {

  constructor(props: IReadySpfxTeams1Props) {
    super(props);

    // set initial state for the component: not loading
    this.state = {
      loading: false,
      Message: '',
      OppInfo: undefined,
      UserInfo: undefined,
    };
  }

  public render(): React.ReactElement<IReadySpfxTeams1Props> {

    let siteTabTitle: string = '';
    let AdditionalTeamInfo: any = [];

    let oppId: string = '';
    let ProjName: string = '';
    let Domain: string = '';
    let PrimaryProduct: string = '';
    let Industry: string = '';

    let UserName: string = '';
    let UPN: string = '';
    let mail: string = '';

    if(this.state.OppInfo){
      oppId = this.state.OppInfo.opportunityId;
      ProjName = this.state.OppInfo.name;
      Domain = this.state.OppInfo.domain;
      PrimaryProduct = this.state.OppInfo.primaryProduct;
      Industry = this.state.OppInfo.industry;
    }

    if(this.state.UserInfo){
      UserName = this.state.UserInfo.Name;
      UPN = this.state.UserInfo.UPN;
      mail = this.state.UserInfo.mail;
    }

    ///[context] get current context of the webpart
    if (this.props._teamContext) {
      siteTabTitle = "We are in the context of following Team: " + this.props._teamContext.teamName;
      AdditionalTeamInfo = [
        <li>Additional Team Info:</li>,
        <li>Channel Name = {this.props._teamContext.channelName}</li>,
        <li>Host Client Type = {this.props._teamContext.hostClientType}</li>,
        <li>Login hint = {this.props._teamContext.loginHint}</li>,
        <li>UPN = {this.props._teamContext.userPrincipalName}</li>
      ];
    }
    else if(this.props.context.pageContext)
    {
      siteTabTitle = "We are in the context of following site: " + this.props.context.pageContext.web.title;
    }
    else
    {
      siteTabTitle = "We are not having any context now!";
    }

    return (
      <div className={ styles.readySpfxTeams1 }>
        <div className={ styles.container }>
              <div className={styles.table} id="results">
              <div className={styles.theader}>
                <div className={styles.table_header}>Attribute</div>
                <div className={styles.table_header}>Value</div>
              </div>
              <div className={styles.table_row}>{strings.OppInfoHeading}
            </div>
              <div className={styles.table_row}>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header One</div>
                  <div className={styles.table_cell}>{strings.OpportunityIdLable}</div>
                </div>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header Two</div>
                  <div className={styles.table_cell}>{oppId}</div>
                </div>
              </div>
              <div className={styles.table_row}>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header One</div>
                  <div className={styles.table_cell}>{strings.DomainLable}</div>
                </div>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header Two</div>
                  <div className={styles.table_cell}>{Domain}</div>
                </div>
              </div>
              <div className={styles.table_row}>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header One</div>
                  <div className={styles.table_cell}>{strings.ProjectNameLable}</div>
                </div>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header Two</div>
                  <div className={styles.table_cell}>{ProjName}</div>
                </div>
              </div>
              <div className={styles.table_row}>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header One</div>
                  <div className={styles.table_cell}>{strings.PrimaryProduct}</div>
                </div>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header Two</div>
                  <div className={styles.table_cell}>{PrimaryProduct}</div>
                </div>
              </div>
              <div className={styles.table_row}>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header One</div>
                  <div className={styles.table_cell}>{strings.Industry}</div>
                </div>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header Two</div>
                  <div className={styles.table_cell}>{Industry}</div>
                </div>
            </div>
            <div className={styles.table_row}>{strings.UserInfoHeading}
            </div>
            <div className={styles.table_row}>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header One</div>
                  <div className={styles.table_cell}>{strings.UserNameLable}</div>
                </div>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header Two</div>
                  <div className={styles.table_cell}>{UserName}</div>
                </div>
            </div>
            <div className={styles.table_row}>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header One</div>
                  <div className={styles.table_cell}>{strings.UPNLable}</div>
                </div>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header Two</div>
                  <div className={styles.table_cell}>{UPN}</div>
                </div>
            </div>
            <div className={styles.table_row}>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header One</div>
                  <div className={styles.table_cell}>
                    <PrimaryButton id="btncreateoppinfo" onClick={this.CreateOppInfo} text={strings.CreateOppInfoBtnTitle} />
                  </div>
                </div>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header Two</div>
                  <div className={styles.table_cell}>{strings.spacer}</div>
                </div>
            </div>
            <div className={styles.table_row}>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header One</div>
                  <div className={styles.table_cell}>
                    <PrimaryButton id="btngetoppinfo" onClick={this.GetOppInfo} text={strings.GetOppInfoBtnTitle} />
                  </div>
                </div>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header Two</div>
                  <div className={styles.table_cell}>{strings.spacer}</div>
                </div>
            </div>
            <div className={styles.table_row}>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header One</div>
                  <div className={styles.table_cell}>
                    <PrimaryButton id="btnupdateoppinfo" onClick={this.UpdateOppInfo} text={strings.UpdateOppInfoBtnTitle} />
                  </div>
                </div>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header Two</div>
                  <div className={styles.table_cell}>{strings.spacer}</div>
                </div>
            </div>
            <div className={styles.table_row}>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header One</div>
                  <div className={styles.table_cell}>
                    <PrimaryButton id="btngetuserinfo" onClick={this.GetUserInfo} text={strings.GetUserInfoBtnTitle} />
                  </div>
                </div>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header Two</div>
                  <div className={styles.table_cell}>{strings.spacer}</div>
                </div>
            </div>
            <div className={styles.table_row}>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header One</div>
                  <div className={styles.table_cell}>
                    {this.state.Message}
                  </div>
                </div>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header Two</div>
                  <div className={styles.table_cell}>{strings.spacer}</div>
                </div>
            </div>
            <div className={styles.table_row}>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header One</div>
                  <div className={styles.table_cell}>
                    {AdditionalTeamInfo}
                  </div>
                </div>
                <div className={styles.table_small}>
                  <div className={styles.table_cell}>Header Two</div>
                  <div className={styles.table_cell}>{strings.spacer}</div>
                </div>
            </div>
          </div>
        </div>
      </div>
    );
  }


  ///[GetbyId] Read from API by Id
  @autobind
  private async GetOppInfo(): Promise<void> {
    console.log("GetOppInfo called");
    this.setState({
      Message: 'getting data...',
      loading: true,
    });

    // create an AadHttpClient object to consume the API
    const aadClient: AadHttpClient = await this.props.context.aadHttpClientFactory.getClient(this.props.functionUri);

    console.log("Created aadClient");

    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Accept', 'application/json');

    const requestOptions: IHttpClientOptions = {
      headers: requestHeaders,
    };

    const httpResponse: HttpClientResponse = await aadClient
      .get(
        this.props.functionUri + "/api/OppInfo/" + this.state.OppInfo.id,
        AadHttpClient.configurations.v1,
        requestOptions
      );

    let responsetext = await httpResponse.text();
    const response: IOppInfo = JSON.parse(responsetext);

    this.setState({
      OppInfo: response,
      Message: "Got OppInfo with id: " + response.id,
      loading: false,
    });
    return;
  }


  ///[CreateId] Post to API
  @autobind
  private async CreateOppInfo(): Promise<void> {
    console.log("CreateOppInfo called");
    this.setState({
      Message: 'Creating Opportunity...',
      loading: true,
      OppInfo: undefined,
    });

    // create an AadHttpClient object to consume the API
    const aadClient: AadHttpClient = await this.props.context.aadHttpClientFactory.getClient(this.props.functionUri);

    console.log("Created aadClient");

    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Accept', 'application/json');

    let bodytxt = JSON.stringify(
        { id: "",
        opportunityId:"7-XDFSDFS3",
        name:"My Project",
        domain:"Apps",
        primaryProduct:"Azure",
        industry:"Banking"
      });

    const requestOptions: IHttpClientOptions = {
      headers: requestHeaders,
      body: bodytxt
    };

    ///[post] use POST method to create object
    const httpResponse: HttpClientResponse = await aadClient
      .post(
        this.props.functionUri + "/api/OppInfo",
        AadHttpClient.configurations.v1,
        requestOptions
      );

    let responsetext = await httpResponse.text();
    const response: IOppInfo = JSON.parse(responsetext);

    this.setState({
      OppInfo: response,
      Message: "Got OppInfo with Id: " + response.id,
      loading: false,
    });
    return;
  }

  ///[UpdatebyId] Update value using PUT method on API
  @autobind
  private async UpdateOppInfo(): Promise<void> {
    console.log("UpdateOppInfo called");
    this.setState({
      Message: 'Update Opportunity...',
      loading: true,
    });

    // create an AadHttpClient object to consume the API
    const aadClient: AadHttpClient = await this.props.context.aadHttpClientFactory.getClient(this.props.functionUri);

    console.log("Created aadClient");

    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Accept', 'application/json');

    let updateInd = "Industry no " + (new Date()).getSeconds();
    this.state.OppInfo.industry = updateInd;
    let bodytxt = JSON.stringify(this.state.OppInfo);

    const requestOptions: IHttpClientOptions = {
      headers: requestHeaders,
      method: "PUT",
      body: bodytxt
    };

    const httpResponse: HttpClientResponse = await aadClient
      .fetch(
        this.props.functionUri + "/api/OppInfo/" + this.state.OppInfo.id,
        AadHttpClient.configurations.v1,
        requestOptions
      );

    let responsetext = await httpResponse.text();
    const response: IOppInfo = JSON.parse(responsetext);

    this.setState({
      OppInfo: response,
      Message: "Got OppInfo with Id: " + response.id,
      loading: false,
    });
    return;
  }


  ///[GraphAPI] Use the Microsoft graph API
  @autobind
  private async GetUserInfo(): Promise<void> {
    console.log("getUserInfo called");
    this.setState({
      Message: 'getting User Info data...',
      loading: true,
      UserInfo: undefined,
    });

    // create an AadHttpClient object to consume the API
    const aadClient: AadHttpClient = await this.props.context.aadHttpClientFactory.getClient("https://graph.microsoft.com");

    console.log("Created aadClient");

    const requestHeaders: Headers = new Headers();
    requestHeaders.append('Accept', 'application/json');

    const requestOptions: IHttpClientOptions = {
      headers: requestHeaders,
    };

    let username = '';
    if(this.props.context.pageContext){
      username = this.props.context.pageContext.user.loginName;
    }

    const httpResponse: HttpClientResponse = await aadClient
      .get(
        `https://graph.microsoft.com/v1.0/users/${username}?$select=displayName,mail,userPrincipalName`,
        AadHttpClient.configurations.v1,
        requestOptions
      );

    let responsetext = await httpResponse.json();
    let user: IUserInfo = undefined;
    if(responsetext){
      user = {
      Name:  responsetext.displayName,
      UPN: responsetext.userPrincipalName,
      mail: responsetext.mail,
    };
  }

    this.setState({
      UserInfo: user,
      Message: "Got User info",
      loading: false,
    });

    return;
  }

}

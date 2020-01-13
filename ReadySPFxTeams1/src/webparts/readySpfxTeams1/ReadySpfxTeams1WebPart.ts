import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version, Log } from '@microsoft/sp-core-library';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-webpart-base';
import { DisplayMode } from '@microsoft/sp-core-library';

import * as strings from 'ReadySpfxTeams1WebPartStrings';
import ReadySpfxTeams1 from './components/ReadySpfxTeams1';
import { IReadySpfxTeams1Props } from './components/IReadySpfxTeams1Props';
import { IReadySpfxTeams1WebPartProps } from './IReadySpfxTeams1WebPartProps';

import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';
import * as microsoftTeams from '@microsoft/teams-js';

export default class ReadySpfxTeams1WebPart extends BaseClientSideWebPart<IReadySpfxTeams1WebPartProps> {

  private _defaultfuncionUri: string;

  private _teamsContext: microsoftTeams.Context;

  private needsConfiguration(): boolean {
    // as long as we don't have the configuration settings
    return (!this.properties.functionUri);
  }


  public render(): void {

    const element: React.ReactElement<IReadySpfxTeams1Props > = React.createElement(
      ReadySpfxTeams1,
      {
        functionUri: this.properties.functionUri,
        needsConfiguration: this.needsConfiguration(),
        context: this.context,
        _teamContext: this._teamsContext,
        configureHandler: () => {
          this.context.propertyPane.open();
        },
        errorHandler: (errorMessage: string) => {
          if (this.displayMode === DisplayMode.Edit) {
            this.context.statusRenderer.renderError(this.domElement, errorMessage);
          } else {
            // nothing to do, if we are not in edit Mode
          }
        }
      }
    );

    ReactDom.render(element, this.domElement);
  }

  ///[OnInit] Use OnInit to set Teams context
  protected onInit(): Promise<any> {
    let retVal: Promise<any> = Promise.resolve();
    this._defaultfuncionUri = this.properties.functionUri;
    if (this.context.microsoftTeams) {
      retVal = new Promise((resolve, reject) => {
        this.context.microsoftTeams.getContext(context => {
          this._teamsContext = context;
          resolve();
        });
      });
    }
    return retVal;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  ///[Prop]Setting up property pane for web-part
  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('functionUri', {
                  label: strings.FunctionUriFieldLabel,
                  value: this.properties.functionUri
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

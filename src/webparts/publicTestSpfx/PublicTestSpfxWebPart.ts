import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PublicTestSpfxWebPartStrings';
import PublicTestSpfx from './components/PublicTestSpfx';
import { IPublicTestSpfxProps } from './components/IPublicTestSpfxProps';
import { IEmailProperties } from '@pnp/sp/sputilities';

import { getSP } from './pnpjsConfig';
import "@pnp/sp/sputilities";
import { spfi, SPFI, SPFx } from "@pnp/sp";

export interface IPublicTestSpfxWebPartProps {
  description: string;
}

export default class PublicTestSpfxWebPart extends BaseClientSideWebPart<IPublicTestSpfxWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IPublicTestSpfxProps> = React.createElement(
      PublicTestSpfx,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }



  //#region Test 1
    protected async onInit(): Promise<void> {
      this._environmentMessage = this._getEnvironmentMessage();
      await super.onInit();

      //Initialize our _sp object that we can then use in other packages without having to pass around the context.
      getSP(this.context);

      const sp = getSP();
      const emailProps: IEmailProperties = {
        // To: [this.userIdDictionnary[memberId].mail],
        To: [""],
        CC: [""],
        // BCC: [""],
        BCC: ["larac@teamwaydemo.fr"],
        Subject: "FlowGetMyTasks;",
        Body: "Test email",
        AdditionalHeaders: {
            "content-type": "text/plain"
        }
      };
      
      try {
        const result = await sp.utility.sendEmail(emailProps);
        debugger;
      } catch (error) {
        debugger;
      }
    }
  //#endregion Test 1

  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

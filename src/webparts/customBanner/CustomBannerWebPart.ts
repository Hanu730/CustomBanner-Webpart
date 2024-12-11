import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { SPHttpClient } from '@microsoft/sp-http';
import * as strings from 'CustomBannerWebPartStrings';
import CustomBanner from './components/CustomBanner';
import { ICustomBannerProps } from './components/ICustomBannerProps';

export interface ICustomBannerWebPartProps {
  description: string;
}

export default class CustomBannerWebPart extends BaseClientSideWebPart<ICustomBannerWebPartProps> {
  private imageUrl:string="";
  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ICustomBannerProps> = React.createElement(
      CustomBanner,
      {
        description: this.properties.description,
        imageurl:this.imageUrl,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    try {
      this.imageUrl = await this.getImageFromList();
    } catch (error) {
      console.error('Error fetching image from list:', error);
    }
    return this._getEnvironmentMessage().then(message => {
      console.log(this._isDarkTheme);
      console.log(this._environmentMessage)
      this._environmentMessage = message;
    });
  }
  private async getImageFromList(): Promise<string> {
    const endpoint = `${this.context.pageContext.web.absoluteUrl}/_api/web/lists/getbytitle('Banner')/items?$select=ImageUrl&$orderby=Created desc`;
    const response = await this.context.spHttpClient.get(endpoint, SPHttpClient.configurations.v1);
    const result = await response.json();
  console.log(result);
    if (result.value.length > 0) {
      return result.value[0].ImageUrl.Url; // Assuming 'ImageColumn' holds the image URL
    }
  
    throw new Error('No images found in the list.');
  }
  


  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
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

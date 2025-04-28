import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'SpfxNavRollupWebPartStrings';
import SpfxNavRollup from './components/SpfxNavRollup';
import { ISpfxNavRollupProps } from './components/ISpfxNavRollupProps';

import { PropertyFieldNumber } from '@pnp/spfx-property-controls/lib/PropertyFieldNumber';
import { PropertyFieldColorPicker, PropertyFieldColorPickerStyle } from '@pnp/spfx-property-controls/lib/PropertyFieldColorPicker';

export interface ISpfxNavRollupWebPartProps {
  description: string;
  QueryUrl: string;
  Width: number;
  Height: number;
  Margin: number;
  Padding: number;
  Background: string;
  Color: string;
  Font: string;
  Border: string;
  BorderRadius: string;
  Gradient: string;
  BoxShadow: string;
  TextAlignment: string;
  VerticalAlignment: string;
}


export default class SpfxNavRollupWebPart extends BaseClientSideWebPart<ISpfxNavRollupWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<ISpfxNavRollupProps> = React.createElement(
      SpfxNavRollup,
      {
        context:this.context,
        QueryUrl: this.properties.QueryUrl,
        Width: this.properties.Width,
        Height: this.properties.Height,
        Margin: this.properties.Margin,
        Padding: this.properties.Padding,
        Background: this.properties.Background,
        Color: this.properties.Color,
        Font: this.properties.Font,
        Border: this.properties.Border,
        BorderRadius: this.properties.BorderRadius,
        Gradient: this.properties.Gradient,
        BoxShadow: this.properties.BoxShadow,
        TextAlignment: this.properties.TextAlignment,
        VerticalAlignment: this.properties.VerticalAlignment,
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
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
                }),
                PropertyPaneTextField('QueryUrl', {
                  label: "QueryUrl",
                  //label: strings.SiteUrlFieldLabel,
                  value: this.properties.QueryUrl,
                  //,                  
                }),
                PropertyFieldNumber("Width", {
                  key: "Width",
                  label: "Width of each Rectangle",
                  description: "Width of each Rectangle",
                  value: this.properties.Width,
                  maxValue: 400,
                  minValue: 50,
                  disabled: false
                }),
                PropertyFieldNumber("Height", {
                  key: "Height",
                  label: "Height of each Rectangle",
                  description: "Height of each Rectangle",
                  value: this.properties.Height,
                  maxValue: 400,
                  minValue: 50,
                  disabled: false
                }),
                PropertyFieldNumber("Margin", {
                  key: "Margin",
                  label: "Margin between each Rectangle",
                  description: "Margin between each Rectangle",
                  value: this.properties.Margin,
                  maxValue: 50,
                  minValue: 1,
                  disabled: false
                }),
                PropertyFieldNumber("Padding", {
                  key: "Padding",
                  label: "Padding within each Rectangle",
                  description: "Padding within each Rectangle",
                  value: this.properties.Padding,
                  maxValue: 50,
                  minValue: 1,
                  disabled: false
                }),
                PropertyFieldColorPicker('Background', {
                  label: 'Background Color',
                  selectedColor: this.properties.Background,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                }),
                PropertyFieldColorPicker('Color', {
                  label: 'Font Color',
                  selectedColor: this.properties.Color,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  disabled: false,
                  alphaSliderHidden: false,
                  style: PropertyFieldColorPickerStyle.Full,
                  iconName: 'Precipitation',
                  key: 'colorFieldId'
                }),
                PropertyPaneTextField('Font', {
                  label: "Font",
                }),
                PropertyPaneTextField('Border', {
                  label: "Border",
                }),
                PropertyPaneTextField('BorderRadius', {
                  label: "BorderRadius",
                }),
                PropertyPaneTextField('Gradient', {
                  label: "Gradient",
                }),
                PropertyPaneTextField('BoxShadow', {
                  label: "BoxShadow",
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

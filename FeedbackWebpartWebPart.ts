import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme , ThemeProvider } from '@microsoft/sp-component-base';

import * as strings from 'FeedbackWebpartWebPartStrings';
import FeedbackWebpart from './components/FeedbackWebpart';
import { IFeedbackWebpartProps } from './components/IFeedbackWebpartProps';
import { property } from 'lodash';
import { ListFieldLabel } from 'FeedbackWebpartWebPartStrings';



export interface IFeedbackWebpartWebPartProps {
  ListTitle: string;
  ListUrl: string;
  QuestionText : string;
}

export default class FeedbackWebpartWebPart extends BaseClientSideWebPart<IFeedbackWebpartWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';
  private themeProvider: ThemeProvider;
  private themeVariant: IReadonlyTheme | undefined;


  public render(): void {
    const element: React.ReactElement<IFeedbackWebpartProps> = React.createElement(
      FeedbackWebpart,
      {
        ListName : this.properties.ListTitle,
        PageName : window.location.href,
        isDarkTheme: this._isDarkTheme,
        userDisplayName: this.context.pageContext.user.displayName,
        context: this.context,
        themeVariant : this.themeVariant,
        listitemid: !!this.context.pageContext.listItem ? this.context.pageContext.listItem.id : null,
        QuestionText: this.properties.QuestionText
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    return super.onInit();
  }

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
                PropertyPaneTextField('ListTitle', {
                  label: strings.ListFieldLabel
                }),
                PropertyPaneTextField('ListUrl',{
                  label: strings.ListUrlLabel,
                 
                  onGetErrorMessage(value:string):string {
                    if(value.length>256){
                      return "URL should be less than 256 character"
                    }
                    if(value.length == 0){
                      return "Please Enter the List Url"
                    }
                    return ""
                  },
                }),
                PropertyPaneTextField('QuestionText',{
                  label: strings.QuestionTextLabel,
                  
                  onGetErrorMessage(value:string):string {
                   
                    if(value.length == 0){
                      return "Please Enter the Question Text "
                    }
                    return ""
                  },
                })

               
              ]
            }
          ]
        }
      ]
    };
  }
}

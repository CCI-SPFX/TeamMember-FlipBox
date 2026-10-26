import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IFlipBoxProps } from './components/IFlipBoxProps';
import FlipBox from './components/FlipBox'; 

export default class FlipBoxWebPart extends BaseClientSideWebPart<IFlipBoxProps> {

  public render(): void {
    const element: React.ReactElement<IFlipBoxProps> = React.createElement(FlipBox, {
      context: this.context,
      listName: this.properties.listName,
      linkedInIconUrl: this.properties.linkedInIconUrl,
      xIconUrl: this.properties.xIconUrl
    });

    ReactDom.render(element, this.domElement);
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
            description: 'Team Grid Web Part Settings'
          },
          groups: [
            {
              groupName: 'List Settings',
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: 'Team Members List Name',
                  description: 'List columns: TeamMember, Description, LinkedinURL, XURL'
                })
              ]
            },
            {
              groupName: 'Social Icons',
              groupFields: [
                PropertyPaneTextField('linkedInIconUrl', {
                  label: 'LinkedIn Icon URL'
                }),
                PropertyPaneTextField('xIconUrl', {
                  label: 'X (Twitter) Icon URL'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
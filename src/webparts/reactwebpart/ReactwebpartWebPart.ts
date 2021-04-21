import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ReactwebpartWebPartStrings';
import Reactwebpart from './components/Reactwebpart';
import { IReactwebpartProps } from './components/IReactwebpartProps';

export interface IReactwebpartWebPartProps {
  userListName: string;
  countryListName: string;
}

export default class ReactwebpartWebPart extends BaseClientSideWebPart<IReactwebpartWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IReactwebpartProps> = React.createElement(
      Reactwebpart,
      {
        userListName: this.properties.userListName,
        countryListName: this.properties.countryListName,
        context: this.context
      }
    );

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
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('userListName', {
                  label : "User List Name",
                  value : "Users"
                }),
                PropertyPaneTextField('countryListName', {
                  label: "Country List Name",
                  value : ""
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

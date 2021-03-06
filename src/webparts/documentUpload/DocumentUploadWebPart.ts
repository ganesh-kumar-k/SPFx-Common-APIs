import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'DocumentUploadWebPartStrings';
import DocumentUpload from './components/DocumentUpload';
import { IDocumentUploadProps } from './components/IDocumentUploadProps';

if (window.location.href.indexOf("/_layouts/15/Workbench.aspx") > -1) {
  console.log("Webpart running locally");
  require('../../styles/workbench.scss');
}

export interface IDocumentUploadWebPartProps {
  description: string;
}

export default class DocumentUploadWebPart extends BaseClientSideWebPart<IDocumentUploadWebPartProps> {

  public render(): void {

    const element: React.ReactElement<IDocumentUploadProps> = React.createElement(
      DocumentUpload,
      {
        description: this.properties.description,
        context: this.context
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  // @ts-ignore
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

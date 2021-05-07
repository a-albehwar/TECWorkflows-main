import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CctvInternalLegalMgrWebPartStrings';
import CctvInternalLegalMgr from './components/CctvInternalLegalMgr';
import { ICctvInternalLegalMgrProps } from './components/ICctvInternalLegalMgrProps';

export interface ICctvInternalLegalMgrWebPartProps {
  description: string;
}

export default class CctvInternalLegalMgrWebPart extends BaseClientSideWebPart<ICctvInternalLegalMgrWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICctvInternalLegalMgrProps> = React.createElement(
      CctvInternalLegalMgr,
      {
        description: this.properties.description,
        context: this.context,
        siteurl: this.context.pageContext.site.absoluteUrl,
        weburl:this.context.pageContext.web.absoluteUrl,
        pagecultureId:this.context.pageContext.cultureInfo.currentUICultureName,
        spHttpClient: this.context.spHttpClient,
      }
    );

    ReactDom.render(element, this.domElement);
  }


 
  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }
  /*
  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }
  */
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

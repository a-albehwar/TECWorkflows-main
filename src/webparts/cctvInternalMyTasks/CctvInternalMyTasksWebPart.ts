import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'CctvInternalMyTasksWebPartStrings';
import CctvInternalMyTasks from './components/CctvInternalMyTasks';
import { ICctvInternalMyTasksProps } from './components/ICctvInternalMyTasksProps';

export interface ICctvInternalMyTasksWebPartProps {
  description: string;
}

export default class CctvInternalMyTasksWebPart extends BaseClientSideWebPart<ICctvInternalMyTasksWebPartProps> {

  public render(): void {
    const element: React.ReactElement<ICctvInternalMyTasksProps> = React.createElement(
      CctvInternalMyTasks,
      {
        description: this.properties.description,
        siteurl: this.context.pageContext.site.absoluteUrl,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

 /* protected get dataVersion(): Version {
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

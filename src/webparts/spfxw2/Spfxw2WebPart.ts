import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'Spfxw2WebPartStrings';
import Spfxw2 from './components/Spfxw2';
import { ISpfxw2Props } from './components/ISpfxw2Props';

export interface ISpfxw2WebPartProps {
  description: string;
}

export default class Spfxw2WebPart extends BaseClientSideWebPart<ISpfxw2WebPartProps> {

  public render(): void {
    const element: React.ReactElement<ISpfxw2Props> = React.createElement(
      Spfxw2,
      {
        description: this.properties.description,
        context:this.context,
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

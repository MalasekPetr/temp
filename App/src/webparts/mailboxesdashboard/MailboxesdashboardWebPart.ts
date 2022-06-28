import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  BaseClientSideWebPart,
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-webpart-base';
import { sp } from '@pnp/sp';
import {
  HttpService
} from '../../services';
import * as strings from 'MailboxesdashboardWebPartStrings';
import { Mailboxesdashboard, IMailboxesdashboardProps } from '../../components';
import { IMailboxesdashboardWebPartProps } from './IMailboxesdashboardWebPartProps';
import { Placeholder } from '@pnp/spfx-controls-react/lib/Placeholder';
import { isUndefined } from 'lodash';

export default class MailboxesdashboardWebPart extends BaseClientSideWebPart<IMailboxesdashboardWebPartProps> {
  private async RenderConfigPlaceholder(): Promise<void> {
    const placeholder: React.ReactElement<unknown> = React.createElement(
      Placeholder,
      {
        iconName: 'Edit',
        iconText: 'Configure your web part',
        description: 'Please configure the web part.',
        buttonLabel: 'Configure',
        onConfigure: this.onConfigure
      }
    )
    ReactDom.render(placeholder, this.domElement)
  }

  public render(): void {
    if (isUndefined(this.properties.backendapi) || (this.properties.backendapi.length < 5)) { // TODO: Add serious validation
      this.RenderConfigPlaceholder();
    } else {
      const element: React.ReactElement<IMailboxesdashboardProps> = React.createElement(
        Mailboxesdashboard,
        {
          webpartprops: this.properties
        }
      );  
      ReactDom.render(element, this.domElement);
    }
  }

  protected async onInit(): Promise<void> {
    return super.onInit().then((): void => {
      HttpService.onInit(this.context.httpClient);
      sp.setup({
        spfxContext: this.context
      });
    });
  }

  private onConfigure = () => {
    this.context.propertyPane.open();
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
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
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('backendapi', {
                  label: strings.BackEndApiFieldLabel
                }),
                PropertyPaneTextField('refreshinterval', {
                  label: strings.RefreshIntervalFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

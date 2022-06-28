import * as React from 'react';
import * as ReactDom from 'react-dom';
import {
  IPropertyPaneConfiguration,
  IPropertyPaneDropdownOption,
  IPropertyPaneField,
  IPropertyPaneGroup,
  IPropertyPanePage,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  BaseClientSideWebPart
} from "@microsoft/sp-webpart-base";
import { sp } from '@pnp/sp';
import {
  ConfigurationService, HttpService,
} from '../../services';
import * as strings from 'MailboxWebPartStrings';
import { Mailbox, IMailboxProps } from '../../components';
import { IMailboxWebPartProps } from './IMailboxWebPartProps';
import { IMailBoxApp, IUser } from '../../models';
import { isUndefined } from 'lodash';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { thProperties } from 'office-ui-fabric-react';

export default class MailboxWebPart extends BaseClientSideWebPart<IMailboxWebPartProps> {
  private lists: IPropertyPaneDropdownOption[] = [];
  private mailboxapp: IMailBoxApp | unknown;
  private list: Array<Record<string, string>>;

  protected async onInit(): Promise<void> {
    return super.onInit().then((): void => {
      HttpService.onInit(this.context.httpClient);
      sp.setup({
        spfxContext: this.context
      });
    });
  }

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

  public async render(): Promise<void> {
    if (isUndefined(this.properties.backendapi) || (this.properties.backendapi.length < 5)) { // TODO: Add serious validation
      this.RenderConfigPlaceholder();
    } else {
      if (isUndefined(this.properties.address) || (this.properties.address.length < 5)) { // TODO: Add serious validation
        this.RenderConfigPlaceholder();
      } else {
        const element: React.ReactElement<IMailboxProps> = React.createElement(
          Mailbox,
          {
            backendapi: this.properties.backendapi,
            name: this.properties.name,
            address: this.properties.address,
            appAddress: this.properties.appAddress,
            spWebBaseUrl: this.properties.spWebBaseUrl,
            spDocLibId: this.properties.spDocLibId,
            spListId: this.properties.spListId,
            members: this.properties.members
          }
        )
        ReactDom.render(element, this.domElement);
      }
    }
  }

  private async GetConfiguration(): Promise<void> {
    const backendapi = this.properties.backendapi.replace(/\/$/, "");
    const configService: ConfigurationService = new ConfigurationService(backendapi);  
    const apps: IMailBoxApp[] = await configService.getMailBoxApps() as IMailBoxApp[];
    apps.forEach(app => {
      this.lists.push({key: app.spDocLibId, text: app.address});
    });
  }

  private async GetMailbox(address: string): Promise<IMailBoxApp> {
    const backendapi = this.properties.backendapi.replace(/\/$/, "");
    const configService: ConfigurationService = new ConfigurationService(backendapi);  
    return await configService.getMailBoxApp(address) as IMailBoxApp;
  }

/*   protected async UpdateConfiguration(): Promise<void> {
    if (!isUndefined(this.properties.members) && (this.properties.members !== '')) {
      await this.UpdateMailBoxApp();
      this.configurationService = new ConfigurationService(this.properties.backendapi);
      await this.configurationService.addOrUpdateMailBoxApp(this.mailboxapp);
    } else {
      this.mailboxapp = undefined;
    }
  } */

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  /* protected onPropertyPaneConfigurationComplete(): void {
    this.UpdateConfiguration();
  } */

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    const pages: IPropertyPanePage[] = [
      {
        header: {
          description: strings.PropertyPaneDescription,
        },
        groups: [
          {
            groupName: strings.BasicGroupName,
            groupFields: [
              PropertyPaneTextField("backendapi", {
                label: strings.BackEndApiFieldLabel,
              }),
            ],
          },
        ],
      },
    ];
    
   const groups: IPropertyPaneGroup = pages[0].groups[0] as IPropertyPaneGroup;
    const groupFields: IPropertyPaneField<any>[] = groups.groupFields;
      if (this.properties.backendapi) {
        this.GetConfiguration();
        groupFields.push(
          PropertyPaneDropdown('address', {
            label: strings.AddressFieldLabel,
            options: this.lists,
            selectedKey: this.properties.address ? this.properties.address : 0,
            disabled: false
          })
        ); 
      }

      if (this.properties.name) {
        groupFields.push(
          PropertyPaneTextField('name', {
            label: strings.NameFieldLabel,
            disabled: true
          }),
        );
      }

      if (this.properties.appAddress) {
        groupFields.push(
          PropertyPaneTextField('appAddress', {
            label: strings.AppAddressFieldLabel,
            disabled: true
          }),
        );
      }

      if (this.properties.spDocLibId) {
        groupFields.push(
          PropertyPaneTextField('spDocLibId', {
            label: strings.SpDocLibIdFieldLabel,
            disabled: true
          }),
        );
      }

      if (this.properties.spListId) {
        groupFields.push(
          PropertyPaneTextField('spListId', {
            label: strings.SpListIdFieldLabel,
            disabled: true
          }),
        );
      }

      if (this.properties.spWebBaseUrl) {
        groupFields.push(
          PropertyPaneTextField('spWebBaseUrl', {
            label: strings.SpWebBaseUrlFieldLabel,
            disabled: true
          }),
        );
      }

      if (this.properties.members) {
        groupFields.push(
          PropertyPaneTextField('members', {
            label: strings.MembersFieldLabel,
            disabled: true
          }),
        );
      }

      const panelConfiguration: IPropertyPaneConfiguration = { pages: pages };
      return panelConfiguration;
  }

  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {
    if (propertyPath === "backendapi" && (newValue !== oldValue)) {
      await this.GetConfiguration();
      this.context.propertyPane.refresh();
    }
    if (propertyPath === "address" && (newValue !== oldValue)) {
      const address = this.lists.filter(i => i.key === newValue);
      const mailbox = await this.GetMailbox(address[0].text);
      this.properties.name = mailbox.name
      this.properties.appAddress = mailbox.appAddress
      this.properties.spWebBaseUrl = mailbox.spWebBaseUrl
      this.properties.spDocLibId = mailbox.spDocLibId
      this.properties.spListId = mailbox.spListId
      this.properties.members = mailbox.users.map(i => i.upn).join(';');
      await this.UpdateMailBoxApp();
      this.context.propertyPane.refresh();
    }
  }

  private onConfigure = async () => {
    this.context.propertyPane.open();
  }

  private async UpdateMailBoxApp(): Promise<void> {
    const users: IUser[] = [];
    const members: string[] = this.properties.members.split(';');
    members.forEach((upn: string) => {
      users.push({
        role: 'Member',
        upn: upn
      });
    });
    this.mailboxapp = {
      backendapi: this.properties.backendapi,
      name: this.properties.name,
      address: this.properties.address,
      appAddress: this.properties.appAddress,
      spWebBaseUrl: this.properties.spWebBaseUrl,
      spDocLibId: this.properties.spDocLibId,
      spListId: this.properties.spListId,
      users: users
    };
  }
}

import { IPropertyPaneConfiguration } from '@microsoft/sp-property-pane';
import { BaseAdaptiveCardExtension } from '@microsoft/sp-adaptive-card-extension-base';
import { CardView } from './cardView/CardView';
import { QuickView } from './quickView/QuickView';
import { BasicCardPropertyPane } from './BasicCardPropertyPane';
import {Web} from "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
export interface IContact {
  Name: string;
  Email: string;
  Mobile: string;
  Landline: string;
  Website: string;
  Address: string;
}
export interface IBasicCardAdaptiveCardExtensionProps {
  title: string;
}

export interface IBasicCardAdaptiveCardExtensionState {
  contactsList:IContact[];
  index:number;
}

const CARD_VIEW_REGISTRY_ID: string = 'BasicCard_CARD_VIEW';
export const QUICK_VIEW_REGISTRY_ID: string = 'BasicCard_QUICK_VIEW';

export default class BasicCardAdaptiveCardExtension extends BaseAdaptiveCardExtension<
  IBasicCardAdaptiveCardExtensionProps,
  IBasicCardAdaptiveCardExtensionState
> {
  private _deferredPropertyPane: BasicCardPropertyPane | undefined;

  public onInit(): Promise<void> {
    this.state = { contactsList:[],index:0};
    this.cardNavigator.register(CARD_VIEW_REGISTRY_ID, () => new CardView());
    this.quickViewNavigator.register(QUICK_VIEW_REGISTRY_ID, () => new QuickView());
    return this.GetListItems().then((Response) => {
      let contacts: IContact[];
      contacts= [];
      Response.map((item:IContact) => {
          contacts.push({
              Name: item.Name,
              Email: item.Email,
              Mobile: item.Mobile,
              Landline: item.Landline,
              Website: item.Website,
              Address: item.Address,
          });
      });
      this.setState({
          contactsList:contacts
      });
      return Promise.resolve();
  })
  }
  protected async GetListItems(): Promise < any > {
    const context = Web("https://7zmht7.sharepoint.com/sites/AddressBook");
    var items = await context.lists.getByTitle("Contacts").items.get();
    if (items.length > 0) {
        return items;
    }
    return null;
}
  protected loadPropertyPaneResources(): Promise<void> {
    return import(
      /* webpackChunkName: 'BasicCard-property-pane'*/
      './BasicCardPropertyPane'
    )
      .then(
        (component) => {
          this._deferredPropertyPane = new component.BasicCardPropertyPane();
        }
      );
  }

  protected renderCard(): string | undefined {
    return CARD_VIEW_REGISTRY_ID;
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return this._deferredPropertyPane?.getPropertyPaneConfiguration();
  }
}


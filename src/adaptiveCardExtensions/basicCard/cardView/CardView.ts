import {
  BaseBasicCardView,
  IBasicCardParameters,
  IExternalLinkCardAction,
  IQuickViewCardAction,
  ICardButton
} from '@microsoft/sp-adaptive-card-extension-base';
import * as strings from 'BasicCardAdaptiveCardExtensionStrings';
import { IBasicCardAdaptiveCardExtensionProps, IBasicCardAdaptiveCardExtensionState, QUICK_VIEW_REGISTRY_ID } from '../BasicCardAdaptiveCardExtension';

export class CardView extends BaseBasicCardView<IBasicCardAdaptiveCardExtensionProps, IBasicCardAdaptiveCardExtensionState> {
  public get cardButtons(): [ICardButton] | [ICardButton, ICardButton] | undefined {
    return [
      {
        title: strings.QuickViewButton,
        action: {
          type: 'QuickView',
          parameters: {
            view: QUICK_VIEW_REGISTRY_ID
          }
        }
      }
    ];
  }

  public get data(): IBasicCardParameters {
    var primaryText;
    if(this.state.contactsList.length > 0){
      primaryText=this.state.contactsList.length+" contacts exists."
    }
    else{
      primaryText="No contact exists."
    }
    console.log(this.properties.title);
    return {
      primaryText: primaryText,
      title: "AddressCard"
    };
  }

  public get onCardSelection(): IQuickViewCardAction | IExternalLinkCardAction | undefined {
    return {
      type: 'ExternalLink',
      parameters: {
        target: 'https://7zmht7.sharepoint.com/sites/AddressBook/SitePages/AddressBook.aspx#/'
      }
    };
  }
}

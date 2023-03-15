import { ISPFxAdaptiveCard, BaseAdaptiveCardView } from '@microsoft/sp-adaptive-card-extension-base';
import { IBasicCardAdaptiveCardExtensionProps, IBasicCardAdaptiveCardExtensionState, IContact } from '../BasicCardAdaptiveCardExtension';

export interface IQuickViewData {
  contactsList:IContact[]
  index:number
}

export class QuickView extends BaseAdaptiveCardView<
  IBasicCardAdaptiveCardExtensionProps,
  IBasicCardAdaptiveCardExtensionState,
  IQuickViewData
> {
  public get data(): IQuickViewData {
    console.log(this.state.contactsList);
    return {
      contactsList:this.state.contactsList,
      index: 0,
    };
  }

  public get template(): ISPFxAdaptiveCard {
    return require('./template/QuickViewTemplate.json');
  }
}
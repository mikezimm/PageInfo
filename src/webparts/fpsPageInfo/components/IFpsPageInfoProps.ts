
import { IFPSBasicToggleSetting, IFPSExpandoAudience, ISupportedHost } from '@mikezimm/npmfunctions/dist/Services/PropPane/FPSInterfaces';

import { IAdvancedPagePropertiesProps } from "./AdvPageProps/components/IAdvancedPagePropertiesProps";
import { IPageNavigatorProps } from "./PageNavigator/IPageNavigatorProps";
import { IPinMeState } from "./PinMe/FPSPinMenu";

export interface IFPSPinMenu {
  domElement: HTMLElement;
  pageLayout: ISupportedHost ;// like SinglePageApp etc... this.context[_pageLayout];
}

export interface IFpsPageInfoProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  pageNavigator: IPageNavigatorProps;

  advPageProps: IAdvancedPagePropertiesProps;

  fpsPinMenu: IFPSPinMenu;

}

export interface IFpsPageInfoState {

  pinState: IPinMeState;

}

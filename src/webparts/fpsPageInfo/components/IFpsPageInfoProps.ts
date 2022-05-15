
import { IFPSBasicToggleSetting, IFPSExpandoAudience, ISupportedHost } from '@mikezimm/npmfunctions/dist/Services/PropPane/FPSInterfaces';

import { WebPartContext } from "@microsoft/sp-webpart-base";

import { IWebpartBannerProps, } from '@mikezimm/npmfunctions/dist/HelpPanel/onNpm/bannerProps';

import { DisplayMode, Version } from '@microsoft/sp-core-library';

import { IWebpartHistory, IWebpartHistoryItem2, } from '@mikezimm/npmfunctions/dist/Services/PropPane/WebPartHistoryInterface';


import { IAdvancedPagePropertiesProps } from "./AdvPageProps/components/IAdvancedPagePropertiesProps";
import { IPageNavigatorProps } from "./PageNavigator/IPageNavigatorProps";
import { IPinMeState } from "@mikezimm/npmfunctions/dist/PinMe/FPSPinMenu";
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as React from 'react';

export interface IFPSPinMenu {
  defPinState: IPinMeState;
  forcePinState: boolean;
  domElement: HTMLElement;
  pageLayout: ISupportedHost ;// like SinglePageApp etc... this.context[_pageLayout];

}

export interface IFpsPageInfoProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  themeVariant: IReadonlyTheme | undefined;

  pageInfoStyle: React.CSSProperties;

  feedbackEmail: string;

  //FPS Banner and Options props
  displayMode: DisplayMode;

  //Environement props
  // pageContext: PageContext;
  context: WebPartContext;
  urlVars: {};

  //Banner related props
  errMessage: any;
  bannerProps: IWebpartBannerProps;

  //ADDED FOR WEBPART HISTORY:  
  webpartHistory: IWebpartHistory;

  pageNavigator: IPageNavigatorProps;

  advPageProps: IAdvancedPagePropertiesProps;

  fpsPinMenu: IFPSPinMenu;


}

export interface IFpsPageInfoState {

  pinState: IPinMeState;

  showDevHeader: boolean;
  lastStateChange: string;

  propsExpanded: boolean;
  tocExpanded: boolean;

  showPropsHelp: boolean;
  
}

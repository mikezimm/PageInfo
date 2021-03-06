

import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";

import { ISupportedHost } from "@mikezimm/npmfunctions/dist/Services/PropPane/FPSInterfaces";

import { IExpandAudiences } from "@mikezimm/npmfunctions/dist/Services/PropPane/FPSOptionsExpando";

import { IWebpartHistory, IWebpartHistoryItem2, } from '@mikezimm/npmfunctions/dist/Services/PropPane/WebPartHistoryInterface';
import { IPinMeState } from "@mikezimm/npmfunctions/dist/PinMe/FPSPinMenu";
import { IMinHeading } from "./components/PageNavigator/IPageNavigatorProps";

import { exportIgnorePropsFPS, importBlockPropsFPS } from '@mikezimm/npmfunctions/dist/WebPartInterfaces/ImportProps';

  //Specific for this web part
  export const exportIgnorePropsThis = [ ];

  export const exportIgnoreProps = [ ...exportIgnorePropsFPS, ...exportIgnorePropsThis  ];

  //These props will not be imported even if they are in one of the change arrays above (fail-safe)
  //This was done so user could not manually insert specific props to over-right fail-safes built in to the webpart

  //Specific for this web part
  export const importBlockPropsThis = [ 'showSomeProps' ];

  export const importBlockProps = [ ...importBlockPropsFPS, ...importBlockPropsThis ];

  //This will be in npmFunctions > Services/PropPane/FPSOptionsExpando in next release.
  //  export type IExpandAudiences = 'Site Admins' | 'Site Owners' | 'Page Editors' | 'WWWone';


  export const changePinMe = [ 'defPinState', 'forcePinState' ];
  export const changeTOC = [ 'showTOC', 'minHeadingToShow' ,'description' , 'TOCTitleField', 'tocExpanded' ];
  export const changeProperties = [ 'showSomeProps', 'showCustomProps' , 'showOOTBProps' , 'showApprovalProps' , 'propsTitleField', 'propsExpanded', 'selectedProperties' ];

  export const changeRelated1 = [ 'related1heading', 'related1showItems' , 'related1isExpanded' , 'related1web' , 'related1listTitle', 'related1restFilter', 'related1linkProp', 'related1displayProp', 'relatedStyle' ];
  export const changeRelated2 = [ 'related2heading', 'related2showItems' , 'related2isExpanded' , 'related2web' , 'related2listTitle', 'related2restFilter', 'related2linkProp', 'related2displayProp' ];
  export const changePageLinks = [ 'pageLinksheading', 'pageLinksshowItems' , 'pageLinksisExpanded' , 'pageLinksweb' , 'pageLinkslistTitle', 'pageLinksrestFilter', 'pageLinkslinkProp', 'pageLinksdisplayProp', 'canvasLinks', 'canvasImgs', 'linkSearchBox' ];

  export const changeWebPartStyles = [ 'h1Style', 'h2Style' ,'h3Style' , 'pageInfoStyle', 'tocStyle', 'propsStyle' ];

export interface IFpsPageInfoWebPartProps {

  defPinState: IPinMeState;
  forcePinState: boolean;
  
  showTOC: boolean;
  minHeadingToShow: IMinHeading;
  description: string;
  TOCTitleField: string;
  tocExpanded: boolean;

  h1Style: string;
  h2Style: string;
  h3Style: string;
  pageInfoStyle: string;
  tocStyle: string;
  propsStyle: string;

  showSomeProps: boolean;
  showCustomProps: boolean;
  showOOTBProps: boolean;
  showApprovalProps: boolean;
  propsExpanded: boolean;

  //Copied from AdvancedPagePropertiesWebPart.ts
  propsTitleField: string;
  selectedProperties: string[];

  feedbackEmail: string;

  relatedStyle: string;

  related1heading: string;
  related1showItems: boolean;
  related1isExpanded: boolean;
  related1web: string;
  related1listTitle: string;
  related1AreFiles: boolean;   // Used to include ServerRedirectedEmbedUrl in fetch for alt-click
  related1restFilter: string;
  related1linkProp: string; // aka FileLeaf to open file name, if empty, will just show the value
  related1displayProp: string;

  related2heading: string;
  related2showItems: boolean;
  related2isExpanded: boolean;
  related2web: string;
  related2listTitle: string;
  related2AreFiles: boolean;   // Used to include ServerRedirectedEmbedUrl in fetch for alt-click
  related2restFilter: string;
  related2linkProp: string; // aka FileLeaf to open file name, if empty, will just show the value
  related2displayProp: string;

  pageLinksheading: string;
  pageLinksshowItems: boolean;
  pageLinksisExpanded: boolean;
  pageLinksweb: string;
  pageLinkslistTitle: string;
  pageLinksrestFilter: string;
  pageLinkslinkProp: string; // aka FileLeaf to open file name, if empty, will just show the value
  pageLinksdisplayProp: string;
  canvasLinks: boolean;
  canvasImgs: boolean;
  ignoreDefaultImages: boolean;
  linkSearchBox: boolean;

  uniqueId: string;
  showBannerGear: boolean; // Not in Prop Pane

  //Needed for Expandoramic and PinMenu
  pageLayout: ISupportedHost ;// like SinglePageApp etc... this.context[_pageLayout];

  //2022-02-17:  Added these for expandoramic mode
  enableExpandoramic: boolean;
  expandoDefault: boolean;
  expandoStyle: any;
  expandoPadding: number;
  expandoAudience: IExpandAudiences;

  // expandAlert: boolean;
  // expandConsole: boolean;
  //2022-02-17:  END additions for expandoramic mode

  // Section 15
  //General settings for Banner Options group
  // export interface IWebpartBannerProps {

  //[ 'showBanner', 'bannerTitle', 'showGoToHome', 'showGoToParent', 'homeParentGearAudience', 'bannerStyleChoice', 'bannerStyle', 'bannerCmdStyle', 'bannerHoverEffect', 'showRepoLinks', 'showExport' ];
  showBanner: boolean;
  bannerTitle: string;

  infoElementChoice: string;
  infoElementText: string;

  showGoToHome: boolean;  //defaults to true
  showGoToParent: boolean;  //defaults to true
  homeParentGearAudience: IExpandAudiences;

  bannerStyleChoice: string;
  bannerStyle: string;
  bannerCmdStyle: string;
  lockStyles: boolean;

  bannerHoverEffect: boolean;
  showRepoLinks: boolean;
  showExport: boolean;

  fpsImportProps: string;

  fullPanelAudience : IExpandAudiences;
  replacePanelHTML : any;  //This is the jsx sent to panel for User controled information (aka what reader will see when clicking 'info' button)

  //These are added for the minimum User Panel component ( which turns into the replacePanelHTML component )
  panelMessageDescription1: string; //
  panelMessageSupport: string;
  panelMessageDocumentation: string;
  panelMessageIfYouStill: string;
  documentationLinkDesc: string;
  documentationLinkUrl: string;
  documentationIsValid: boolean;
  supportContacts: IPropertyFieldGroupOrPerson[];

  //ADDED FOR WEBPART HISTORY:  
  webpartHistory: IWebpartHistory;


  showTricks: boolean;

  // }

  //Section 16 - FPS Options group
  searchShow: boolean;
  fpsPageStyle: string;
  fpsContainerMaxWidth: string;
  quickLaunchHide: boolean;

  //FPS Options part II
  pageHeaderHide: boolean;
  allSectionMaxWidthEnable: boolean;
  allSectionMaxWidth: number;
  allSectionMarginEnable: boolean;
  allSectionMargin: number;
  toolBarHide: boolean;

}

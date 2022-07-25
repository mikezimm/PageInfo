

import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";

import { ISupportedHost } from "@mikezimm/npmfunctions/dist/Services/PropPane/FPSInterfaces";

import { IExpandAudiences } from "@mikezimm/npmfunctions/dist/Services/PropPane/FPSOptionsExpando";

import { IWebpartHistory, IWebpartHistoryItem2, } from '@mikezimm/npmfunctions/dist/Services/PropPane/WebPartHistoryInterface';
import { IPinMeState } from "@mikezimm/npmfunctions/dist/PinMe/FPSPinMenu";
import { IMinHeading } from "./components/PageNavigator/IPageNavigatorProps";

import { exportIgnorePropsFPS, importBlockPropsFPS } from '@mikezimm/npmfunctions/dist/WebPartInterfaces/ImportProps';
import { IMinWPBannerProps } from "@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/BannerSetup";
import { IMinRelatedWPProps, IMinPageLinksProps } from "@mikezimm/npmfunctions/dist/RelatedItems/IRelatedWebPartProps";

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

export interface IFpsPageInfoWebPartProps extends IMinWPBannerProps, IMinRelatedWPProps, IMinPageLinksProps {

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

  //2022-07-25:  Moved all related and pageLinks interfaces to npmFunctions where rest of code actually is.

  // relatedStyle: string;

  // related1heading: string;
  // related1showItems: boolean;
  // related1isExpanded: boolean;
  // related1web: string;
  // related1listTitle: string;
  // related1AreFiles: boolean;   // Used to include ServerRedirectedEmbedUrl in fetch for alt-click
  // related1restFilter: string;
  // related1linkProp: string; // aka FileLeaf to open file name, if empty, will just show the value
  // related1displayProp: string;

  // related2heading: string;
  // related2showItems: boolean;
  // related2isExpanded: boolean;
  // related2web: string;
  // related2listTitle: string;
  // related2AreFiles: boolean;   // Used to include ServerRedirectedEmbedUrl in fetch for alt-click
  // related2restFilter: string;
  // related2linkProp: string; // aka FileLeaf to open file name, if empty, will just show the value
  // related2displayProp: string;

  // pageLinksheading: string;
  // pageLinksshowItems: boolean;
  // pageLinksisExpanded: boolean;
  // pageLinksweb: string;
  // pageLinkslistTitle: string;
  // pageLinksrestFilter: string;
  // pageLinkslinkProp: string; // aka FileLeaf to open file name, if empty, will just show the value
  // pageLinksdisplayProp: string;

  // canvasLinks: boolean;
  // canvasImgs: boolean;
  // ignoreDefaultImages: boolean;
  // linkSearchBox: boolean;

}

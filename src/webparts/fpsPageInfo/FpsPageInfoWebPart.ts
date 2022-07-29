import * as React from 'react';
import * as ReactDom from 'react-dom';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  IPropertyPaneDropdownOption,
  IPropertyPaneField,
  PropertyPaneHorizontalRule,
  PropertyPaneDropdown,
  IPropertyPaneDropdownProps,
  PropertyPaneButton,
  PropertyPaneButtonType,
  PropertyPaneToggle,
  IPropertyPaneGroup,
} from '@microsoft/sp-property-pane';


import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";

import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import {   
  ThemeProvider,
  ThemeChangedEventArgs,
  IReadonlyTheme } from '@microsoft/sp-component-base';

import { INavLink } from 'office-ui-fabric-react/lib/Nav';

import { sp, Views, IViews, ISite } from "@pnp/sp/presets/all";

//Copied from AdvancedPagePropertiesWebPart.ts
import * as _lodashAPP from 'lodash';

import { PropertyFieldPeoplePicker, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

import { setPageFormatting, } from '@mikezimm/npmfunctions/dist/Services/DOM/FPSFormatFunctions';

import { IFPSPage, } from '@mikezimm/npmfunctions/dist/Services/DOM/FPSInterfaces';
import { createFPSWindowProps, initializeFPSSection, initializeFPSPage, webpartInstance, initializeMinimalStyle } from '@mikezimm/npmfunctions/dist/Services/DOM/FPSDocument';
import { IFPSWindowProps, IFPSSection, IFPSSectionStyle } from '@mikezimm/npmfunctions/dist/Services/DOM/FPSInterfaces';
import { setSectionStyles } from '@mikezimm/npmfunctions/dist/Services/DOM/setAllSectionStyles';
import { minimizeHeader } from '@mikezimm/npmfunctions/dist/Services/DOM/minimzeHeader';
import { minimizeToolbar } from '@mikezimm/npmfunctions/dist/Services/DOM/minimzeToolbar';
import { minimizeQuickLaunch } from '@mikezimm/npmfunctions/dist/Services/DOM/quickLaunch';
import { applyHeadingCSS } from '@mikezimm/npmfunctions/dist/HeadingCSS/FPSHeadingFunctions';
import { renderCustomStyles } from '@mikezimm/npmfunctions/dist/WebPartFunctions/MainWebPartStyleFunctions';

import { replaceHandleBars } from '@mikezimm/npmfunctions/dist/Services/Strings/handleBars';

// import { FPSOptionsGroupBasic, FPSBanner2Group, FPSOptionsGroupAdvanced } from '@mikezimm/npmfunctions/dist/Services/PropPane/FPSOptionsGroup2';
import { FPSOptionsGroupBasic, FPSBanner3Group, FPSOptionsGroupAdvanced } from '@mikezimm/npmfunctions/dist/Services/PropPane/FPSOptionsGroup3';
import { FPSBanner3BasicGroup,FPSBanner3NavGroup, FPSBanner3ThemeGroup } from '@mikezimm/npmfunctions/dist/Services/PropPane/FPSOptionsGroup3';
import { FPSBanner3VisHelpGroup } from '@mikezimm/npmfunctions/dist/CoreFPS/FPSOptionsGroupVisHelp';
import { FPSPinMePropsGroup } from '@mikezimm/npmfunctions/dist/PinMe/FPSOptionsGroupPinMe';
import { buildRelatedItemsPropsGroup } from '@mikezimm/npmfunctions/dist/RelatedItems/RelatedItemsPropGroup';

import { FPSOptionsExpando, expandAudienceChoicesAll } from '@mikezimm/npmfunctions/dist/Services/PropPane/FPSOptionsExpando'; //expandAudienceChoicesAll

import { WebPartInfoGroup, JSON_Edit_Link } from '@mikezimm/npmfunctions/dist/Services/PropPane/zReusablePropPane';


import { _LinkIsValid } from '@mikezimm/npmfunctions/dist/Links/AllLinks';
import * as links from '@mikezimm/npmfunctions/dist/Links/LinksRepos';

import { importProps, FPSImportPropsGroup } from '@mikezimm/npmfunctions/dist/Services/PropPane/ImportFunctions';

import { sortStringArray, sortObjectArrayByStringKey, sortNumberArray, sortObjectArrayByNumberKey, sortKeysByOtherKey 
} from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import { IBuildBannerSettings , buildBannerProps, IMinWPBannerProps } from '@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/BannerSetup';

import { buildExportProps, buildFPSAnalyticsProps } from './CoreFPS/BuildExportProps';

import { setExpandoRamicMode } from '@mikezimm/npmfunctions/dist/Services/DOM/FPSExpandoramic';
import { getUrlVars } from '@mikezimm/npmfunctions/dist/Services/Logging/LogFunctions';

//encodeDecodeString(this.props.libraryPicker, 'decode')
import { encodeDecodeString, } from "@mikezimm/npmfunctions/dist/Services/Strings/urlServices";

import { verifyAudienceVsUser } from '@mikezimm/npmfunctions/dist/Services/Users/CheckPermissions';

import { bannerThemes, bannerThemeKeys, makeCSSPropPaneString, createBannerStyleStr, createBannerStyleObj } from '@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/defaults';

import { IRepoLinks } from '@mikezimm/npmfunctions/dist/Links/CreateLinks';

import { IWebpartHistory, IWebpartHistoryItem2 } from '@mikezimm/npmfunctions/dist/Services/PropPane/WebPartHistory/Interface';
import { createWebpartHistory, ITrimThis, updateWebpartHistoryV2, upgradeV1History } from '@mikezimm/npmfunctions/dist/Services/PropPane/WebPartHistory/Functions';
import { getWebPartHistoryOnInit } from '@mikezimm/npmfunctions/dist/Services/PropPane/WebPartHistory/OnInit';

import { saveAnalytics3 } from '@mikezimm/npmfunctions/dist/Services/Analytics/analytics2';
import { IZLoadAnalytics, IZSentAnalytics, } from '@mikezimm/npmfunctions/dist/Services/Analytics/interfaces';

import { getSiteInfo, getWebInfoIncludingUnique } from '@mikezimm/npmfunctions/dist/Services/Sites/getSiteInfo';
import { IFPSUser } from '@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces';
import { getFPSUser } from '@mikezimm/npmfunctions/dist/Services/Users/FPSUser';

// import { startPerformInit, startPerformOp, updatePerformanceEnd } from './components/Performance/functions';
// import { IPerformanceOp, ILoadPerformanceALVFM, IHistoryPerformance } from './components/Performance/IPerformance';
import { IWebpartBannerProps } from '@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/bannerProps';

import { ISupportedHost } from "@mikezimm/npmfunctions/dist/Services/PropPane/FPSInterfaces";

export const repoLink: IRepoLinks = links.gitRepoPageInfoSmall;

require('@mikezimm/npmfunctions/dist/Services/PropPane/GrayPropPaneAccordions.css');
require('@mikezimm/npmfunctions/dist/PinMe/FPSPinMe.css');
require('@mikezimm/npmfunctions/dist/HeadingCSS/FPSHeadings.css');
require('@mikezimm/npmfunctions/dist/PropPaneHelp/PropPanelHelp.css');

import { SPService } from '../../Service/SPService';

import { visitorPanelInfo } from '@mikezimm/npmfunctions/dist/CoreFPS/VisitorPanelComponent';
import * as strings from 'FpsPageInfoWebPartStrings';
import FpsPageInfo from './components/FpsPageInfo';
import { IFpsPageInfoProps } from './components/IFpsPageInfoProps';


import { Log } from './components/AdvPageProps/utilities/Log';
import { IFpsPageInfoWebPartProps } from './IFpsPageInfoWebPartProps';
import { exportIgnoreProps, importBlockProps, } from './IFpsPageInfoWebPartProps';
import { createStyleFromString } from '@mikezimm/npmfunctions/dist/Services/PropPane/StringToReactCSS';
import { FPSApplyHeadingCSS, FPSApplyTagCSSAndStyles, FPSApplyHeadingStyle } from './components/HeadingCSS/FPSTagFunctions';
import { HTMLRegEx, IHTMLRegExKeys } from '../../Service/htmlTags';
import { css } from 'office-ui-fabric-react';
import { PreConfiguredProps } from './CoreFPS/PreConfiguredSettings';
import { getThisSitesPreConfigProps, IConfigurationProp, ISitePreConfigProps, IPreConfigSettings, IAllPreConfigSettings } from '@mikezimm/npmfunctions/dist/PropPaneHelp/PreConfigFunctions';
import { applyPresetCollectionDefaults } from '@mikezimm/npmfunctions/dist/PropPaneHelp/ApplyPresets';
import { IRelatedItemsProps, IRelatedKey } from '@mikezimm/npmfunctions/dist/RelatedItems/IRelatedItemsProps';
import { buildPagePropertiesGroup } from './PropPaneGroups/PageProps';
import { buildImageLinksGroup } from './PropPaneGroups/ImageLinks';

import { buildTOCGroup } from './PropPaneGroups/TOC';
import { buildPageInfoStylesGroup } from './PropPaneGroups/PageInfoStyles';
import { updateBannerStyles } from './CoreFPS/BannerStyleFunctions';


export default class FpsPageInfoWebPart extends BaseClientSideWebPart<IFpsPageInfoWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  private h1 = {  classes: [],  css: '',  };
  private h2 = {  classes: [],  css: '',  };
  private h3 = {  classes: [],  css: '',  };

  //Common FPS variables

  private sitePresets : ISitePreConfigProps = null;

  private _unqiueId;
  private validDocsContacts: string = '';

  private trickyApp = 'FPS PageInfo';
  private wpInstanceID: any = webpartInstance( this.trickyApp );

  private FPSUser: IFPSUser = null;

  private urlParameters: any = {};

  //For FPS options
  private fpsPageDone: boolean = false;
  private fpsPageArray: any[] = null;
  private minQuickLaunch: boolean = false;
  private minHideToolbar: boolean = false;

  //For FPS Banner
  private forceBanner = true ;
  private modifyBannerTitle = true ;
  private modifyBannerStyle = true ;

  private  expandoDefault = false;
  private filesList: any = [];

  private exitPropPaneChanged = false;

  private expandoErrorObj = {

  };

  //ADDED FOR WEBPART HISTORY:  
  private thisHistoryInstance: IWebpartHistoryItem2 = null;

  private importErrorMessage = '';
    
  // private performance : ILoadPerformanceALVFM = null;
  private bannerProps: IWebpartBannerProps = null;

  private beAReader: boolean = false; //2022-04-07:  Intent of this is a one-time per instance to 'become a reader' level user.  aka, hide banner buttons that reader won't see

  //Added from react-page-navigator component
  private anchorLinks: INavLink[] = [];

  //Copied from AdvancedPagePropertiesWebPart.ts
  private availableProperties: IPropertyPaneDropdownOption[] = [];
  private _themeProvider: ThemeProvider;
  private _themeVariant: IReadonlyTheme | undefined;

  private analyticsWasExecuted: boolean = false;

  private _handleThemeChangedEvent(args: ThemeChangedEventArgs): void {
    this._themeVariant = args.theme;
    this.render();
  }


  /***
 *     .d88b.  d8b   db      d888888b d8b   db d888888b d888888b 
 *    .8P  Y8. 888o  88        `88'   888o  88   `88'   `~~88~~' 
 *    88    88 88V8o 88         88    88V8o 88    88       88    
 *    88    88 88 V8o88         88    88 V8o88    88       88    
 *    `8b  d8' 88  V888        .88.   88  V888   .88.      88    
 *     `Y88P'  VP   V8P      Y888888P VP   V8P Y888888P    YP    
 *                                                               
 *                                                               
 */

  protected onInit(): Promise<void> {
    this._environmentMessage = this._getEnvironmentMessage();

    // Consume the new ThemeProvider service
    this._themeProvider = this.context.serviceScope.consume(ThemeProvider.serviceKey);

    // If it exists, get the theme variant
    this._themeVariant = this._themeProvider.tryGetTheme();

    // Register a handler to be notified if the theme variant changes
    this._themeProvider.themeChangedEvent.add(this, this._handleThemeChangedEvent);

    return super.onInit().then(async _ => {

      // other init code may be present

      let mess = 'onInit - ONINIT: ' + new Date().toLocaleTimeString();
      console.log(mess);

      //https://stackoverflow.com/questions/52010321/sharepoint-online-full-width-page
      if ( window.location.href &&  
        window.location.href.toLowerCase().indexOf("layouts/15/workbench.aspx") > 0 ) {
          
        if (document.getElementById("workbenchPageContent")) {
          document.getElementById("workbenchPageContent").style.maxWidth = "none";
        }
      } 

      //console.log('window.location',window.location);
      sp.setup({
        spfxContext: this.context
      });


      /***
     *     .d88b.  d8b   db      d888888b d8b   db d888888b d888888b      d8888b. db   db  .d8b.  .d8888. d88888b      .d888b. 
     *    .8P  Y8. 888o  88        `88'   888o  88   `88'   `~~88~~'      88  `8D 88   88 d8' `8b 88'  YP 88'          VP  `8D 
     *    88    88 88V8o 88         88    88V8o 88    88       88         88oodD' 88ooo88 88ooo88 `8bo.   88ooooo         odD' 
     *    88    88 88 V8o88         88    88 V8o88    88       88         88~~~   88~~~88 88~~~88   `Y8b. 88~~~~~       .88'   
     *    `8b  d8' 88  V888        .88.   88  V888   .88.      88         88      88   88 88   88 db   8D 88.          j88.    
     *     `Y88P'  VP   V8P      Y888888P VP   V8P Y888888P    YP         88      YP   YP YP   YP `8888Y' Y88888P      888888D 
     *                                                                                                                         
     *                                                                                                                         
     */

      //NEED TO APPLY THIS HERE as well as follow-up in render for it to not visibly change
      // this.presetCollectionDefaults();
      this.sitePresets = applyPresetCollectionDefaults( this.sitePresets, PreConfiguredProps, this.properties, this.context.pageContext.web.serverRelativeUrl ) ;
      applyHeadingCSS( this.properties );

      this.properties.pageLayout =  this.context['_pageLayoutType']?this.context['_pageLayoutType'] : this.context['_pageLayoutType'];



      // DEFAULTS SECTION:  Performance   <<< ================================================================
      // this.performance = startPerformInit( this.displayMode, false );

      // DEFAULTS SECTION:  FPSUser


      // (property) BaseClientSideWebPart<IAlvFinManWebPartProps>.context: WebPartContext
      // {@inheritDoc @microsoft/sp-component-base#BaseComponent.context}

      // Argument of type 'import("C:/Users/dev/Documents/GitHub/ALVFinMan7/node_modules/@microsoft/sp-webpart-base/dist/index-internal").WebPartContext' is not assignable to parameter of type 'import("C:/Users/dev/Documents/GitHub/ALVFinMan7/node_modules/@mikezimm/npmfunctions/node_modules/@microsoft/sp-webpart-base/dist/index-internal").WebPartContext'.
      //   Types have separate declarations of a private property '_domElement'.ts(2345)
      //Typed this.context as any to remove above error
      this.FPSUser = getFPSUser( this.context as any, links.trickyEmails, this.trickyApp ) ;
      console.log( 'FPSUser: ', this.FPSUser );


      // // DEFAULTS SECTION:  Expandoramic   <<< ================================================================
      // this.expandoDefault = this.properties.expandoDefault === true && this.properties.enableExpandoramic === true && this.displayMode === DisplayMode.Read ? true : false;
      // if ( this.urlParameters.Mode === 'Edit' ) { this.expandoDefault = false; }
      // let expandoStyle: any = {};

      // //2022-04-07:  Could use the function for parsing JSON for this... check npmFunctions
      // try {
      //   expandoStyle = JSON.parse( this.properties.expandoStyle );

      // } catch(e) {
      //   console.log('Unable to expandoStyle: ', this.properties.expandoStyle);
      // }

      // let padding = this.properties.expandoPadding ? this.properties.expandoPadding : 20;
      // setExpandoRamicMode( this.context.domElement, this.expandoDefault, expandoStyle,  false, false, padding, this.properties.pageLayout  );

      // Moved to ForceEverywhere in src\PropPaneHelp\PreConfiguredConstants.ts
      // this.properties.showRepoLinks = false;
      // this.properties.showExport = false;
      // this.properties.enableExpandoramic = false; //Always hide expandoramic for PageInfo Web Part
      // this.properties.showBanner = true; //Always show banner

      // DEFAULTS SECTION:  Banner   <<< ================================================================
      //This updates unlocks styles only when bannerStyleChoice === custom.  Rest are locked in the ui.

      updateBannerStyles( this.properties, this.context.pageContext.site.serverRelativeUrl , 'corpDark1' );

      if ( this.properties.bannerStyleChoice === 'custom' ) { 
        this.properties.lockStyles = false ; 

      } else { this.properties.lockStyles = true; }

      // if ( this.properties.bannerHoverEffect === undefined ) { this.properties.bannerHoverEffect = false; }

      let defBannerTheme = 'corpDark1';
      if ( this.context.pageContext.site.serverRelativeUrl.toLowerCase().indexOf( '/sites/lifenet') === 0 ) {
          defBannerTheme = 'corpWhite1'; }

      if ( !this.properties.bannerStyle ) { this.properties.bannerStyle = createBannerStyleStr( defBannerTheme, 'banner') ; }

      if ( !this.properties.bannerCmdStyle ) { 

        //Adjust the default size down compared to PinMe buttons which are primary functions in the web part
        let bannerCmdStyle = createBannerStyleStr( defBannerTheme, 'cmd').replace('"fontSize":20,', '"fontSize":16,') ;
        bannerCmdStyle = bannerCmdStyle.replace('"marginRight":"9px"', '"marginRight":"0px"') ;
        bannerCmdStyle = bannerCmdStyle.replace('"padding":"7px"', '"padding":"7px 4px"') ;

        this.properties.bannerCmdStyle = bannerCmdStyle;

       }

      // DEFAULTS SECTION:  Panel   <<< ================================================================
      // Moved to PresetFPSBanner in src\PropPaneHelp\PreConfiguredConstants.ts
      // if ( !this.properties.fullPanelAudience || this.properties.fullPanelAudience.length === 0 ) {
      //   this.properties.fullPanelAudience = 'Page Editors';
      // }
      // if ( !this.properties.documentationLinkDesc || this.properties.documentationLinkDesc.length === 0 ) {
      //   this.properties.documentationLinkDesc = 'Documentation';
      // }


      // DEFAULTS SECTION:  webPartHistory   <<< ================================================================
      //Preset this on existing installations
      // if ( this.properties.forceReloadScripts === undefined || this.properties.forceReloadScripts === null ) {
      //   this.properties.forceReloadScripts = false;
      // }
      //ADDED FOR WEBPART HISTORY:  This sets the webpartHistory
      // this.thisHistoryInstance = createWebpartHistory( 'onInit' , 'new', this.context.pageContext.user.displayName );
      // let priorHistory : IWebpartHistoryItem2[] = this.properties.webpartHistory ? upgradeV1History( this.properties.webpartHistory ).history : [];
      // this.properties.webpartHistory = {
      //   thisInstance: this.thisHistoryInstance,
      //   history: priorHistory,
      // };

      this.properties.webpartHistory = getWebPartHistoryOnInit( this.context.pageContext.user.displayName, this.properties.webpartHistory );

      //Added from react-page-navigator component
      let tags : IHTMLRegExKeys = 'h14';
      if ( this.properties.minHeadingToShow === 'h2' ) {
        tags = 'h13';
      } else if ( this.properties.minHeadingToShow === 'h1' ) {
        tags = 'h12';
      }

      this.anchorLinks = await SPService.GetAnchorLinks( this.context, tags );

      //Moved these to src\webparts\fpsPageInfo\CoreFPS\PreConfiguredSettings.ts
      // if ( this.properties.propsExpanded === undefined || this.properties.propsExpanded === null ) { this.properties.propsExpanded = true; }
      // if ( this.properties.propsTitleField === undefined || this.properties.propsTitleField === null ) { this.properties.propsTitleField = strings.bannerTitle; }

      //Have to insure selectedProperties always is an array from AdvancedPagePropertiesWebPart.ts
      // if ( !this.properties.selectedProperties ) { this.properties.selectedProperties = []; }

      renderCustomStyles( this as any, this.domElement, this.properties, false );

    });
  }

  public render(): void {

    /***
 *    d8888b. d88888b d8b   db d8888b. d88888b d8888b. 
 *    88  `8D 88'     888o  88 88  `8D 88'     88  `8D 
 *    88oobY' 88ooooo 88V8o 88 88   88 88ooooo 88oobY' 
 *    88`8b   88~~~~~ 88 V8o88 88   88 88~~~~~ 88`8b   
 *    88 `88. 88.     88  V888 88  .8D 88.     88 `88. 
 *    88   YD Y88888P VP   V8P Y8888D' Y88888P 88   YD 
 *                                                     
 *                                                     
 */


  /***
   *    d8888b. d88888b d8b   db d8888b. d88888b d8888b.       .o88b.  .d8b.  db      db      .d8888. 
   *    88  `8D 88'     888o  88 88  `8D 88'     88  `8D      d8P  Y8 d8' `8b 88      88      88'  YP 
   *    88oobY' 88ooooo 88V8o 88 88   88 88ooooo 88oobY'      8P      88ooo88 88      88      `8bo.   
   *    88`8b   88~~~~~ 88 V8o88 88   88 88~~~~~ 88`8b        8b      88~~~88 88      88        `Y8b. 
   *    88 `88. 88.     88  V888 88  .8D 88.     88 `88.      Y8b  d8 88   88 88booo. 88booo. db   8D 
   *    88   YD Y88888P VP   V8P Y8888D' Y88888P 88   YD       `Y88P' YP   YP Y88888P Y88888P `8888Y' 
   *                                                                                                  
   *           Source:   PivotTiles 1.5.2.6                                                                                
   */
    renderCustomStyles( this as any, this.domElement, this.properties, false );

    this.properties.showSomeProps = this.properties.showOOTBProps === true || this.properties.showCustomProps === true || this.properties.showApprovalProps === true  ? true : false;

    //Preset infoElement to question mark circle for this particular web part if it's not specificed - due to pin icon being important and usage in pinned location
    // if ( !this.properties.infoElementChoice ) { this.properties.infoElementChoice = 'IconName=Unknown'; }  //May not be normally needed if in the presets
    // if ( !this.properties.infoElementText ) { this.properties.infoElementText = 'Question mark circle'; }  //May not be normally needed if in the presets

    this._unqiueId = this.context.instanceId;

    // quickRefresh is used for SecureScript for when caching html file.  <<< ================================================================
    let renderAsReader = this.displayMode === DisplayMode.Read && this.beAReader === true ? true : false;

    let errMessage = '';
    this.validDocsContacts = ''; //This may no longer be needed if links below are commented out.

    // if ( this.properties.documentationIsValid !== true ) { errMessage += ' Invalid Support Doc Link: ' + ( this.properties.documentationLinkUrl ? this.properties.documentationLinkUrl : 'Empty.  ' ) ; this.validDocsContacts += 'DocLink,'; }
    // if ( !this.properties.supportContacts || this.properties.supportContacts.length < 1 ) { errMessage += ' Need valid Support Contacts' ; this.validDocsContacts += 'Contacts,'; }

    let errorObjArray :  any[] =[];

    /***
      *    d8888b.  .d8b.  d8b   db d8b   db d88888b d8888b. 
      *    88  `8D d8' `8b 888o  88 888o  88 88'     88  `8D 
      *    88oooY' 88ooo88 88V8o 88 88V8o 88 88ooooo 88oobY' 
      *    88~~~b. 88~~~88 88 V8o88 88 V8o88 88~~~~~ 88`8b   
      *    88   8D 88   88 88  V888 88  V888 88.     88 `88. 
      *    Y8888P' YP   YP VP   V8P VP   V8P Y88888P 88   YD 
      *                                                      
      *                                                      
      */

    let replacePanelWarning = `Anyone with lower permissions than '${this.properties.fullPanelAudience}' will ONLY see this content in panel`;

    console.log('mainWebPart: buildBannerSettings ~ 387',   );

    let buildBannerSettings : IBuildBannerSettings = {

      FPSUser: this.FPSUser,
      //this. related info
      context: this.context as any,
      clientWidth: ( this.domElement.clientWidth - ( this.displayMode === DisplayMode.Edit ? 250 : 0) ),
      exportProps: buildExportProps( this.properties, this.wpInstanceID, this.context.pageContext.web.serverRelativeUrl ),

      //Webpart related info
      panelTitle: 'FPS Page Info',
      modifyBannerTitle: this.modifyBannerTitle,
      repoLinks: repoLink,

      //Hard-coded Banner settings on webpart itself
      forceBanner: this.forceBanner,
      earyAccess: false,
      wideToggle: false,
      expandAlert: false,
      expandConsole: false,

      replacePanelWarning: replacePanelWarning,
      //Error info
      errMessage: errMessage,
      errorObjArray: errorObjArray, //In the case of Pivot Tiles, this is manualLinks[],
      expandoErrorObj: this.expandoErrorObj,

      beAUser: renderAsReader,
      showBeAUserIcon: false,

    };

    // console.log('mainWebPart: showTricks ~ 322',   );
    let showTricks: any = false;
    links.trickyEmails.map( getsTricks => {
      if ( this.context.pageContext.user && this.context.pageContext.user.loginName && this.context.pageContext.user.loginName.toLowerCase().indexOf( getsTricks ) > -1 ) { 
        showTricks = true ;
        this.properties.showRepoLinks = true; //Always show these users repo links
      }
      } );

    // console.log('mainWebPart: verifyAudienceVsUser ~ 341',   );
    this.properties.showBannerGear = verifyAudienceVsUser( this.FPSUser , showTricks, this.properties.homeParentGearAudience, null, renderAsReader );

    let bannerSetup = buildBannerProps( this.properties , this.FPSUser, buildBannerSettings, showTricks, renderAsReader, this.displayMode );
    if ( !this.properties.bannerTitle || this.properties.bannerTitle === '' ) { 
      if ( this.properties.defPinState !== 'normal' ) {
        bannerSetup.bannerProps.title = strings.bannerTitle ;
      } else {
        bannerSetup.bannerProps.title = 'hide' ;
      }
    }

    errMessage = bannerSetup.errMessage;
    this.bannerProps = bannerSetup.bannerProps;
    let expandoErrorObj = bannerSetup.errorObjArray;

    this.bannerProps.enableExpandoramic = false; //Hard code this option for FPS PageInfo web part only because of PinMe option

    //Add this to force a title because when pinned by default, users may not know it's there.
    if ( this.properties.forcePinState === true && this.properties.defPinState !== 'normal' ) {
      if ( !this.properties.bannerTitle || this.properties.bannerTitle.length < 3 ) { this.bannerProps.title = 'Page Contents' ; }
    }
    // if ( this.bannerProps.showBeAUserIcon === true ) { this.bannerProps.beAUserFunction = this.beAUserFunction.bind(this); }

    // console.log('mainWebPart: visitorPanelInfo ~ 405',   );
    this.properties.replacePanelHTML = visitorPanelInfo( this.properties, repoLink, '', '' );

    this.bannerProps.replacePanelHTML = this.properties.replacePanelHTML;

    const OOTBProps = this.properties.showOOTBProps === true ? ['ID', 'Modified', 'Editor' , 'Created', 'Author' ] : [];
    const ApprovalProps = []; //this.properties.showApprovalProps === true ? ['ID', 'Created', 'Modified'] : [];
    const CustomProps = this.properties.showCustomProps ?  this.properties.selectedProperties : [];

    // let selectedProperties = [ ...CustomProps, ...OOTBProps, ...ApprovalProps ];
    let selectedProperties = [ ...CustomProps ];

    // else if ( this.props.styleString ) { bannerStyle = createStyleFromString( this.props.styleString, { background: 'green' }, 'bannerStyle in banner/component.tsx ~ 81' ); }

    let pageInfoStyle: React.CSSProperties = createStyleFromString( this.properties.pageInfoStyle, { paddingBottom: '20px', background: '#d3d3d3' }, 'FPSPageInfoWP in ~ 511' );
    let tocStyle: React.CSSProperties = createStyleFromString( this.properties.tocStyle, null, 'FPSPageInfoWP in ~ 512' );
    let propsStyle: React.CSSProperties = createStyleFromString( this.properties.propsStyle, null, 'FPSPageInfoWP in ~ 513' );
    let relatedItemsStyle: React.CSSProperties = createStyleFromString( this.properties.relatedStyle, null, 'FPSPageInfoWP in ~ 514' );

    const element: React.ReactElement<IFpsPageInfoProps> = React.createElement(
      FpsPageInfo,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        themeVariant: this._themeVariant,

        //Environement props
        // pageContext: this.context.pageContext, //This can be found in the bannerProps now
        context: this.context,
        urlVars: getUrlVars(),
        displayMode: this.displayMode,

        saveLoadAnalytics: this.saveLoadAnalytics.bind(this),

        pageInfoStyle: pageInfoStyle,

        //Banner related props
        errMessage: 'any',
        bannerProps: this.bannerProps,
        webpartHistory: this.properties.webpartHistory,

        sitePresets: this.sitePresets,

        pageNavigator:   {
          minHeadingToShow: this.properties.minHeadingToShow,
          showTOC: this.properties.showTOC,
          description: this.properties.TOCTitleField,
          anchorLinks: this.anchorLinks,
          themeVariant: this._themeVariant,
          tocExpanded: this.properties.tocExpanded,
          tocStyle: tocStyle,
        },

        advPageProps: {
          defaultExpanded: this.properties.propsExpanded,
          showSomeProps: this.properties.showSomeProps,
          showOOTBProps: this.properties.showOOTBProps,
          context: this.context,
          title: this.properties.propsTitleField,
          selectedProperties: selectedProperties,
          themeVariant: this._themeVariant,
          propsStyle: propsStyle,
        },

        relatedItemsProps1: {
          parentKey: 'related1',
          heading: replaceHandleBars( this.properties.related1heading, this.context ),
          showItems: this.properties.related1showItems,
          fetchInfo: {
            web: this.properties.related1web.toLowerCase() === 'current' ? this.context.pageContext.web.serverRelativeUrl : replaceHandleBars( this.properties.related1web, this.context ),
            listTitle: replaceHandleBars( this.properties.related1listTitle, this.context ),
            restFilter: replaceHandleBars( this.properties.related1restFilter, this.context ),
            itemsAreFiles: this.properties.related1AreFiles, // aka FileLeaf to open file name, if empty, will just show the value
            linkProp: this.properties.related1linkProp, // aka FileLeaf to open file name, if empty, will just show the value
            displayProp: this.properties.related1displayProp,
          },
          isExpanded: this.properties.related1isExpanded,
          themeVariant: this._themeVariant,
          itemsStyle: relatedItemsStyle,
        },

        relatedItemsProps2: {
          parentKey: 'related2',
          heading: replaceHandleBars( this.properties.related2heading, this.context ) ,
          showItems: this.properties.related2showItems,
          fetchInfo: {
            web: this.properties.related2web.toLowerCase() === 'current' ? this.context.pageContext.web.serverRelativeUrl : replaceHandleBars( this.properties.related2web, this.context ),
            listTitle: replaceHandleBars( this.properties.related2listTitle, this.context ),
            restFilter: replaceHandleBars( this.properties.related2restFilter, this.context ),
            itemsAreFiles: this.properties.related2AreFiles, // aka FileLeaf to open file name, if empty, will just show the value
            linkProp: this.properties.related2linkProp, // aka FileLeaf to open file name, if empty, will just show the value
            displayProp: this.properties.related2displayProp,
          },
          isExpanded: this.properties.related2isExpanded,
          themeVariant: this._themeVariant,
          itemsStyle: relatedItemsStyle,
        },

        pageLinks: {
          parentKey: 'pageLinks',
          heading: replaceHandleBars( this.properties.pageLinksheading, this.context ) ,
          showItems: this.properties.pageLinksshowItems,
          fetchInfo: {
            web: this.properties.pageLinksweb.toLowerCase() === 'current' ? this.context.pageContext.web.serverRelativeUrl : replaceHandleBars( this.properties.pageLinksweb, this.context ),
            listTitle: replaceHandleBars( this.properties.pageLinkslistTitle, this.context ),
            restFilter: replaceHandleBars( this.properties.pageLinksrestFilter, this.context ),
            linkProp: this.properties.pageLinkslinkProp, // aka FileLeaf to open file name, if empty, will just show the value
            itemsAreFiles: false,
            displayProp: this.properties.pageLinksdisplayProp,
            canvasLinks: this.properties.canvasLinks,
            canvasImgs: this.properties.canvasImgs,
            ignoreDefaultImages: this.properties.ignoreDefaultImages,
          },
          linkSearchBox: this.properties.linkSearchBox,
          isExpanded: this.properties.pageLinksisExpanded,
          themeVariant: this._themeVariant,
          itemsStyle: relatedItemsStyle,
        },

        fpsPinMenu: {
          defPinState: this.properties.defPinState,
          forcePinState: this.properties.forcePinState,
          domElement: this.context.domElement,
          pageLayout: this.properties.pageLayout,
        }
        
      }
    );

    ReactDom.render(element, this.domElement);

  }


  // private replaceHandleBars( str: string , context: WebPartContext ) {
  // If needed, see npmFunctions\src\Services\Strings\handleBars.ts
  // }

  private beAUserFunction() {
    console.log('beAUserFunction:',   );
    if ( this.displayMode === DisplayMode.Edit ) {
      alert("'Be a regular user' mode is only available while viewing the page.  \n\nOnce you are out of Edit mode, please refresh the page (CTRL-F5) to reload the web part.");

    } else {
      this.beAReader = this.beAReader === true ? false : true;
      this.render();
    }

  }

  /***
   *    d888888b db   db d88888b .88b  d88. d88888b 
   *    `~~88~~' 88   88 88'     88'YbdP`88 88'     
   *       88    88ooo88 88ooooo 88  88  88 88ooooo 
   *       88    88~~~88 88~~~~~ 88  88  88 88~~~~~ 
   *       88    88   88 88.     88  88  88 88.     
   *       YP    YP   YP Y88888P YP  YP  YP Y88888P 
   *                                                
   *                                                
   */
  
  private _getEnvironmentMessage(): string {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams
      return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
    }

    return this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment;
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;
    this.domElement.style.setProperty('--bodyText', semanticColors.bodyText);
    this.domElement.style.setProperty('--link', semanticColors.link);
    this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered);

  }

  //Copied from AdvancedPagePropertiesWebPart.ts
  private async getPageProperties(): Promise<void> {
    Log.Write("Getting Site Page fields...");
    const list = await sp.web.lists.ensureSitePagesLibrary();
    const fi = await list.fields();

    this.availableProperties = [];
    Log.Write(`${fi.length.toString()} fields retrieved!`);
    fi.forEach((f) => {
      if (!f.FromBaseType && !f.Hidden && f.SchemaXml.indexOf("ShowInListSettings=\"FALSE\"") === -1
          && f.TypeAsString !== "Boolean" && f.TypeAsString !== "Note") {
        const internalFieldName = f.InternalName == "LinkTitle" ? "Title" : f.InternalName;
        this.availableProperties.push({ key: internalFieldName, text: f.Title });
        Log.Write(f.TypeAsString);
      }
    });
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  //Copied from AdvancedPagePropertiesWebPart.ts
  protected onAddButtonClick (value: any) {
    this.properties.selectedProperties.push(this.availableProperties[0].key.toString());
  }

  //Copied from AdvancedPagePropertiesWebPart.ts
  protected onDeleteButtonClick (value: any) {
    Log.Write(value.toString());
    var removed = this.properties.selectedProperties.splice(value, 1);
    Log.Write(`${removed[0]} removed.`);
  }


    
  /***
 *    d8888b. d8888b.  .d88b.  d8888b.      d8888b.  .d8b.  d8b   db d88888b       .o88b. db   db  .d8b.  d8b   db  d888b  d88888b 
 *    88  `8D 88  `8D .8P  Y8. 88  `8D      88  `8D d8' `8b 888o  88 88'          d8P  Y8 88   88 d8' `8b 888o  88 88' Y8b 88'     
 *    88oodD' 88oobY' 88    88 88oodD'      88oodD' 88ooo88 88V8o 88 88ooooo      8P      88ooo88 88ooo88 88V8o 88 88      88ooooo 
 *    88~~~   88`8b   88    88 88~~~        88~~~   88~~~88 88 V8o88 88~~~~~      8b      88~~~88 88~~~88 88 V8o88 88  ooo 88~~~~~ 
 *    88      88 `88. `8b  d8' 88           88      88   88 88  V888 88.          Y8b  d8 88   88 88   88 88  V888 88. ~8~ 88.     
 *    88      88   YD  `Y88P'  88           88      YP   YP VP   V8P Y88888P       `Y88P' YP   YP YP   YP VP   V8P  Y888P  Y88888P 
 *                                                                                                                                 
 *                                                                                                                                 
 */

  //Copied from AdvancedPagePropertiesWebPart.ts
  // protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
  protected async onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any) {
    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);

    if (propertyPath.indexOf("selectedProperty") >= 0) {
      Log.Write('Selected Property identified');
      let index: number = _lodashAPP.toInteger(propertyPath.replace("selectedProperty", ""));
      this.properties.selectedProperties[index] = newValue;
    }


    if ( propertyPath === 'documentationLinkUrl' || propertyPath === 'fpsImportProps' ) {
      this.properties.documentationIsValid = await _LinkIsValid( newValue ) === "" ? true : false;
      console.log( `${ newValue ? newValue : 'Empty' } Docs Link ${ this.properties.documentationIsValid === true ? ' IS ' : ' IS NOT ' } Valid `);

    } else {
      if ( !this.properties.documentationIsValid ) { this.properties.documentationIsValid = false; }

    }

    this.properties.webpartHistory = updateWebpartHistoryV2( this.properties.webpartHistory , propertyPath , newValue, this.context.pageContext.user.displayName, [], [] );

    if ( propertyPath === 'fpsImportProps' ) {

      if ( this.exitPropPaneChanged === true ) {//Added to prevent re-running this function on import.  Just want re-render. )
        this.exitPropPaneChanged = false;  //Added to prevent re-running this function on import.  Just want re-render.

      } else {
        let result = importProps( this.properties, newValue, [], importBlockProps );

        this.importErrorMessage = result.errMessage;
        if ( result.importError === false ) {
          this.properties.fpsImportProps = '';
          this.context.propertyPane.refresh();
        }
        this.exitPropPaneChanged = true;  //Added to prevent re-running this function on import.  Just want re-render.
        this.onPropertyPaneConfigurationStart();
        // this.render();
      }

    } else if ( propertyPath === 'bannerStyle' || propertyPath === 'bannerCmdStyle' )  {
      this.properties[ propertyPath ] = newValue;
      this.context.propertyPane.refresh();

    } else if (propertyPath === 'bannerStyleChoice')  {
      // bannerThemes, bannerThemeKeys, makeCSSPropPaneString

      if ( newValue === 'custom' ) {
        this.properties.lockStyles = false;

      } else if ( newValue === 'lock') {
        this.properties.lockStyles = true;

      } else {
        this.properties.lockStyles = true;

        let bannerStyle = createBannerStyleStr( newValue, 'banner' );
        
        //Adjust the default size down compared to PinMe buttons which are primary functions in the web part
        let bannerCmdStyle = createBannerStyleStr( newValue, 'cmd' ).replace('"fontSize":20,', '"fontSize":16,');  
        bannerCmdStyle = bannerCmdStyle.replace('"marginRight":"9px"', '"marginRight":"0px"') ;
        bannerCmdStyle = bannerCmdStyle.replace('"padding":"7px"', '"padding":"7px 4px"') ;


        this.properties.bannerStyle = bannerStyle;
        this.properties.bannerCmdStyle = bannerCmdStyle;

        //Reset main web part styles to defaults
        this.properties.pageInfoStyle = '"paddingBottom":"20px","backgroundColor":"#d3d3d3"';
        this.properties.tocStyle = "";
        this.properties.propsStyle = "";
        this.properties.h1Style = "";
        this.properties.h2Style = "";
        this.properties.h3Style = "";

      }

    }

    this.context.propertyPane.refresh();

    this.render();

  }

  //Copied from AdvancedPagePropertiesWebPart.ts
  protected async onPropertyPaneConfigurationStart(): Promise<void> {
    Log.Write(`onPropertyPaneConfigurationStart`);
    await this.getPageProperties();
    this.context.propertyPane.refresh();
  }


  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {

    Log.Write(`getPropertyPaneConfiguration`);

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          displayGroupsAsAccordion: true, //DONT FORGET THIS IF PROP PANE GROUPS DO NOT EXPAND
          groups: [
            WebPartInfoGroup( links.gitRepoPageInfoSmall, 'Best TOC and Page Info available :)' ),
            FPSPinMePropsGroup, //End this group  
            buildTOCGroup( this.properties ), //End this group
            buildPagePropertiesGroup( this.properties, this.availableProperties, this ),
            buildRelatedItemsPropsGroup( this.properties, 'related1' ),
            buildRelatedItemsPropsGroup( this.properties, 'related2' ),
            buildImageLinksGroup( this.properties ),
            buildPageInfoStylesGroup( this.properties, this.modifyBannerStyle ), //End this group
            FPSBanner3VisHelpGroup( this.context, this.onPropertyPaneFieldChanged, this.properties ),
            FPSBanner3BasicGroup( this.forceBanner , this.modifyBannerTitle, this.properties.showBanner, this.properties.infoElementChoice === 'Text' ? true : false, true ),
            FPSBanner3NavGroup(), 
            FPSBanner3ThemeGroup( this.modifyBannerStyle, this.properties.showBanner, this.properties.lockStyles, ),
            FPSOptionsGroupBasic( false, true, true, true, this.properties.allSectionMaxWidthEnable, true, this.properties.allSectionMarginEnable, true ), // this group
            // FPSOptionsExpando( this.properties.enableExpandoramic, this.properties.enableExpandoramic,null, null ),
  
            FPSImportPropsGroup, // this group
          ]
        }
      ]
    };
  }

  /***
 *    d88888b d8888b. .d8888.       .d88b.  d8888b. d888888b d888888b  .d88b.  d8b   db .d8888. 
 *    88'     88  `8D 88'  YP      .8P  Y8. 88  `8D `~~88~~'   `88'   .8P  Y8. 888o  88 88'  YP 
 *    88ooo   88oodD' `8bo.        88    88 88oodD'    88       88    88    88 88V8o 88 `8bo.   
 *    88~~~   88~~~     `Y8b.      88    88 88~~~      88       88    88    88 88 V8o88   `Y8b. 
 *    88      88      db   8D      `8b  d8' 88         88      .88.   `8b  d8' 88  V888 db   8D 
 *    YP      88      `8888Y'       `Y88P'  88         YP    Y888888P  `Y88P'  VP   V8P `8888Y' 
 *                                                                                              
 *                                                                                              
 */

  // private renderCustomStyles( doHeadings: boolean = true ) {

  //   if ( doHeadings === true ) this.applyHeadingCSS();

  //   //Used with FPS Options Functions
  //   this.setQuickLaunch( this.properties.quickLaunchHide );
  //   minimizeHeader( document, this.properties.pageHeaderHide, false, true );
  //   this.setThisPageFormatting( this.properties.fpsPageStyle );
  //   this.setToolbar( this.properties.toolBarHide );
  //   this.updateSectionStyles( );
  // }

  // /**
  //  * Used with FPS Options Functions
  //  * @param quickLaunchHide 
  //  */
  //  private setQuickLaunch( quickLaunchHide: boolean ) {
  //   if ( quickLaunchHide === true && this.minQuickLaunch === false ) {
  //     minimizeQuickLaunch( document , quickLaunchHide );
  //     this.minQuickLaunch = true;
  //   }
  // }

  // /**
  //  * Used with FPS Options Functions
  //  * @param quickLaunchHide 
  //  */
  // private setToolbar( hideToolbar: boolean ) {
  //     if(this.displayMode == DisplayMode.Read && this.urlParameters.tool !== 'true' ){
  //       let value = hideToolbar === true ? 'none' : null;
  //       let toolBarStyle: IFPSSectionStyle = initializeMinimalStyle( 'Miminze Toolbar', this.wpInstanceID, 'display', value );
  //       minimizeToolbar( document, toolBarStyle, false, true );
  //       this.minHideToolbar = true;
  //     }
  // }

  // /**
  //  * Used with FPS Options Functions
  //  * @param fpsPageStyle 
  //  */
  // private setThisPageFormatting( fpsPageStyle: string ) {

  //   let fpsPage: IFPSPage = initializeFPSPage( this.wpInstanceID, this.fpsPageDone, fpsPageStyle, this.fpsPageArray  );
  //   fpsPage = setPageFormatting( this.domElement, fpsPage );
  //   this.fpsPageArray = fpsPage.Array;
  //   this.fpsPageDone = fpsPage.do;

  // }


  // private updateSectionStyles( ) {

  //   let allSectionMaxWidth = this.properties.allSectionMaxWidthEnable !== true ? null : this.properties.allSectionMaxWidth;
  //   let allSectionMargin = this.properties.allSectionMarginEnable !== true ? null : this.properties.allSectionMargin;
  //   let sectionStyles = initializeFPSSection( this.wpInstanceID, allSectionMaxWidth, allSectionMargin,  );

  //   setSectionStyles( document, sectionStyles, true, true );

  // }

  // private applyHeadingCSS() {

  //   if ( this.properties.h1Style ) {
  //     let pieces = this.properties.h1Style.split(';');
  //     let classes = [];
  //     let cssStyles = [];
  //     pieces.map( piece => {
  //       piece = piece.trim();
  //       if ( piece.indexOf('.') === 0 ) { classes.push( piece.replace('.','') ) ; } else { cssStyles.push( piece ) ; }
  //     });

  //     if ( cssStyles.length > 0 || classes.length > 0 ) FPSApplyTagCSSAndStyles( HTMLRegEx.h2, cssStyles.join( ';' ) , classes, true, false, );
  //   }

  //   if ( this.properties.h2Style ) {
  //     let pieces = this.properties.h2Style.split(';');
  //     let classes = [];
  //     let cssStyles = [];
  //     pieces.map( piece => {
  //       piece = piece.trim();
  //       if ( piece.indexOf('.') === 0 ) { classes.push( piece.replace('.','') ) ; } else { cssStyles.push( piece ) ; }
  //     });

  //     if ( cssStyles.length > 0 || classes.length > 0 ) FPSApplyTagCSSAndStyles( HTMLRegEx.h3, cssStyles.join( ';' ) , classes, true, false, );

  //   }

  //   if ( this.properties.h3Style ) {
  //     let pieces = this.properties.h3Style.split(';');
  //     let classes = [];
  //     let cssStyles = [];
  //     pieces.map( piece => {
  //       piece = piece.trim();
  //       if ( piece.indexOf('.') === 0 ) { classes.push( piece.replace('.','') ) ; } else { cssStyles.push( piece ) ; }
  //     });

  //     if ( cssStyles.length > 0 || classes.length > 0 ) FPSApplyTagCSSAndStyles( HTMLRegEx.h4, cssStyles.join( ';' ) , classes, true, false, );

  //   }
  // }


/***
 *     .d8b.  d8b   db  .d8b.  db      db    db d888888b d888888b  .o88b. .d8888. 
 *    d8' `8b 888o  88 d8' `8b 88      `8b  d8' `~~88~~'   `88'   d8P  Y8 88'  YP 
 *    88ooo88 88V8o 88 88ooo88 88       `8bd8'     88       88    8P      `8bo.   
 *    88~~~88 88 V8o88 88~~~88 88         88       88       88    8b        `Y8b. 
 *    88   88 88  V888 88   88 88booo.    88       88      .88.   Y8b  d8 db   8D 
 *    YP   YP VP   V8P YP   YP Y88888P    YP       YP    Y888888P  `Y88P' `8888Y' 
 *                                                                                
 *                                                                                
 */
  private async saveLoadAnalytics( Title: string, Result: string, ) {

    if ( this.analyticsWasExecuted === true ) {
      console.log('saved view info already');

    } else {

      // Do not save anlytics while in Edit Mode... only after save and page reloads
      if ( this.displayMode === DisplayMode.Edit ) { return; }

      let loadProperties: IZLoadAnalytics = {
        SiteID: this.context.pageContext.site.id['_guid'] as any,  //Current site collection ID for easy filtering in large list
        WebID:  this.context.pageContext.web.id['_guid'] as any,  //Current web ID for easy filtering in large list
        SiteTitle:  this.context.pageContext.web.title as any, //Web Title
        TargetSite:  this.context.pageContext.web.serverRelativeUrl,  //Saved as link column.  Displayed as Relative Url
        ListID:  `${this.context.pageContext.list.id}`,  //Current list ID for easy filtering in large list
        ListTitle:  this.context.pageContext.list.title,
        TargetList: `${this.context.pageContext.web.serverRelativeUrl}`,  //Saved as link column.  Displayed as Relative Url

      };

      let zzzRichText1Obj = null;
      let zzzRichText2Obj = null;
      let zzzRichText3Obj = null;

      console.log( 'zzzRichText1Obj:', zzzRichText1Obj);
      console.log( 'zzzRichText2Obj:', zzzRichText2Obj);
      console.log( 'zzzRichText3Obj:', zzzRichText3Obj);

      let zzzRichText1 = null;
      let zzzRichText2 = null;
      let zzzRichText3 = null;

      //This will get rid of all the escaped characters in the summary (since it's all numbers)
      // let zzzRichText3 = ''; //JSON.stringify( fetchInfo.summary ).replace('\\','');
      //This will get rid of the leading and trailing quotes which have to be removed to make it real json object
      // zzzRichText3 = zzzRichText3.slice(1, zzzRichText3.length - 1);

      if ( zzzRichText1Obj ) { zzzRichText1 = JSON.stringify( zzzRichText1Obj ); }
      if ( zzzRichText2Obj ) { zzzRichText2 = JSON.stringify( zzzRichText2Obj ); }
      if ( zzzRichText3Obj ) { zzzRichText3 = JSON.stringify( zzzRichText3Obj ); }

      console.log('zzzRichText1 length:', zzzRichText1 ? zzzRichText1.length : 0 );
      console.log('zzzRichText2 length:', zzzRichText2 ? zzzRichText2.length : 0 );
      console.log('zzzRichText3 length:', zzzRichText3 ? zzzRichText3.length : 0 );

      let FPSProps = null;
      let FPSPropsObj = buildFPSAnalyticsProps( this.properties, this.wpInstanceID, this.context.pageContext.web.serverRelativeUrl );
      FPSProps = JSON.stringify( FPSPropsObj );

      let saveObject: IZSentAnalytics = {
        loadProperties: loadProperties,

        Title: Title,  //General Label used to identify what analytics you are saving:  such as Web Permissions or List Permissions.

        Result: Result,  //Success or Error

        zzzText1: `${ this.properties.defPinState } - ${ this.properties.forcePinState ===  true ? 'forced' : '' }`,

        zzzText2: `${ this.properties.showTOC } - ${  ( this.properties.tocExpanded  ===  true ? 'expanded' : '' ) } - ${  !this.properties.TOCTitleField ? 'Empty Title' : this.properties.TOCTitleField }`,
        zzzText3: `${ this.properties.minHeadingToShow }`,

        zzzText4: `${ this.properties.showSomeProps } - ${ this.properties.propsExpanded  ===  true ? 'expanded' : 'collapsed' } -${ !this.properties.propsTitleField ? 'Empty Title' : this.properties.propsTitleField }`,
        zzzText5: `${ this.properties.showOOTBProps } - ${ this.properties.showCustomProps } - ${ this.properties.showApprovalProps }}`,

        //Info1 in some webparts.  Simple category defining results.   Like Unique / Inherited / Collection
        zzzText6: `${   this.properties.selectedProperties ? this.properties.selectedProperties.join('; ') : '' }`, //Info2 in some webparts.  Phrase describing important details such as "Time to check old Permissions: 86 snaps / 353ms"

        // zzzNumber1: fetchInfo.fetchTime,
        // zzzNumber2: fetchInfo.regexTime,
        // zzzNumber3: fetchInfo.Block.length,
        // zzzNumber4: fetchInfo.Warn.length,
        // zzzNumber5: fetchInfo.Verify.length,
        // zzzNumber6: fetchInfo.Secure.length,
        // zzzNumber7: fetchInfo.js.length,

        zzzRichText1: zzzRichText1,  //Used to store JSON objects for later use, will be stringified
        zzzRichText2: zzzRichText2,
        zzzRichText3: zzzRichText3,

        FPSProps: FPSProps,

      };

      saveAnalytics3( strings.analyticsWeb , `${strings.analyticsList}` , saveObject, true );

      this.analyticsWasExecuted = true;
      console.log('saved view info');

    }

  }

}

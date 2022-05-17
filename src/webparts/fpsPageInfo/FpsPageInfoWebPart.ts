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


} from '@microsoft/sp-property-pane';


import { IPropertyFieldGroupOrPerson } from "@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker";

import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
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

// import { FPSOptionsGroupBasic, FPSBanner2Group, FPSOptionsGroupAdvanced } from '@mikezimm/npmfunctions/dist/Services/PropPane/FPSOptionsGroup2';
import { FPSOptionsGroupBasic, FPSBanner3Group, FPSOptionsGroupAdvanced } from '@mikezimm/npmfunctions/dist/Services/PropPane/FPSOptionsGroup3';
import { FPSBanner3BasicGroup,FPSBanner3NavGroup, FPSBanner3ThemeGroup } from '@mikezimm/npmfunctions/dist/Services/PropPane/FPSOptionsGroup3';

import { FPSOptionsExpando, expandAudienceChoicesAll } from '@mikezimm/npmfunctions/dist/Services/PropPane/FPSOptionsExpando'; //expandAudienceChoicesAll

import { WebPartInfoGroup, JSON_Edit_Link } from '@mikezimm/npmfunctions/dist/Services/PropPane/zReusablePropPane';


import { _LinkIsValid } from '@mikezimm/npmfunctions/dist/Links/AllLinks';
import * as links from '@mikezimm/npmfunctions/dist/Links/LinksRepos';

import { importProps, } from '@mikezimm/npmfunctions/dist/Services/PropPane/ImportFunctions';

import { sortStringArray, sortObjectArrayByStringKey, sortNumberArray, sortObjectArrayByNumberKey, sortKeysByOtherKey 
} from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';

import { IBuildBannerSettings , buildBannerProps, IMinWPBannerProps } from '@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/BannerSetup';

import { buildExportProps } from './BuildExportProps';

import { setExpandoRamicMode } from '@mikezimm/npmfunctions/dist/Services/DOM/FPSExpandoramic';
import { getUrlVars } from '@mikezimm/npmfunctions/dist/Services/Logging/LogFunctions';

//encodeDecodeString(this.props.libraryPicker, 'decode')
import { encodeDecodeString, } from "@mikezimm/npmfunctions/dist/Services/Strings/urlServices";

import { verifyAudienceVsUser } from '@mikezimm/npmfunctions/dist/Services/Users/CheckPermissions';

import { bannerThemes, bannerThemeKeys, makeCSSPropPaneString, createBannerStyleStr, createBannerStyleObj } from '@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/defaults';

import { IRepoLinks } from '@mikezimm/npmfunctions/dist/Links/CreateLinks';

import { IWebpartHistory, IWebpartHistoryItem2 } from '@mikezimm/npmfunctions/dist/Services/PropPane/WebPartHistoryInterface';
import { createWebpartHistory, ITrimThis, updateWebpartHistory, upgradeV1History } from '@mikezimm/npmfunctions/dist/Services/PropPane/WebPartHistoryFunctions';

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
require('./components/HeadingCSS/FPSHeadings.css');
require('@mikezimm/npmfunctions/dist/PropPaneHelp/PropPanelHelp.css');

import { SPService } from '../../Service/SPService';

import { visitorPanelInfo } from './components/VisitorPanel/PanelComponent';
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
import { getThisSitesPreConfigProps, IConfigurationProp, ISitePreConfigProps, PreConfiguredPrpos } from './PreConfiguredSettings';


//export type IMinHeading = 'h3' | 'h2' | 'h1' ;
export const MinHeadingOptions = [
  { index: 0, key: 'h3', text: "h3" },
  { index: 1, key: 'h2', text: "h2" },
  { index: 2, key: 'h1', text: "h1" },
];

//export type IPinMeState = 'normal' | 'pinFull' | 'pinMini';
export const PinMeLocations = [
  { index: 0, key: 'normal', text: "normal" },
  { index: 1, key: 'pinFull', text: "Pin Expanded" },
  { index: 2, key: 'pinMini', text: "Pin Collapsed" },
];

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
      this.presetCollectionDefaults();
      this.applyHeadingCSS();

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
      this.properties.showRepoLinks = false;
      this.properties.showExport = false;
      this.properties.enableExpandoramic = false; //Always hide expandoramic for PageInfo Web Part
      this.properties.showBanner = true; //Always show banner

      // DEFAULTS SECTION:  Banner   <<< ================================================================
      //This updates unlocks styles only when bannerStyleChoice === custom.  Rest are locked in the ui.
      if ( this.properties.bannerStyleChoice === 'custom' ) { this.properties.lockStyles = false ; } else { this.properties.lockStyles = true; }

      if ( this.properties.bannerHoverEffect === undefined ) { this.properties.bannerHoverEffect = false; }

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
      if ( !this.properties.fullPanelAudience || this.properties.fullPanelAudience.length === 0 ) {
        this.properties.fullPanelAudience = 'Page Editors';
      }
      if ( !this.properties.documentationLinkDesc || this.properties.documentationLinkDesc.length === 0 ) {
        this.properties.documentationLinkDesc = 'Documentation';
      }


      // DEFAULTS SECTION:  webPartHistory   <<< ================================================================
      //Preset this on existing installations
      // if ( this.properties.forceReloadScripts === undefined || this.properties.forceReloadScripts === null ) {
      //   this.properties.forceReloadScripts = false;
      // }
      //ADDED FOR WEBPART HISTORY:  This sets the webpartHistory
      this.thisHistoryInstance = createWebpartHistory( 'onInit' , 'new', this.context.pageContext.user.displayName );
      let priorHistory : IWebpartHistoryItem2[] = this.properties.webpartHistory ? upgradeV1History( this.properties.webpartHistory ).history : [];
      this.properties.webpartHistory = {
        thisInstance: this.thisHistoryInstance,
        history: priorHistory,
      };

      //Added from react-page-navigator component
      let tags : IHTMLRegExKeys = 'h14';
      if ( this.properties.minHeadingToShow === 'h2' ) {
        tags = 'h13';
      } else if ( this.properties.minHeadingToShow === 'h1' ) {
        tags = 'h12';
      }

      this.anchorLinks = await SPService.GetAnchorLinks( this.context, tags );

      if ( this.properties.propsExpanded === undefined || this.properties.propsExpanded === null ) { this.properties.propsExpanded = true; }
      if ( this.properties.propsTitleField === undefined || this.properties.propsTitleField === null ) { this.properties.propsTitleField = strings.bannerTitle; }

      //Have to insure selectedProperties always is an array from AdvancedPagePropertiesWebPart.ts
      // if ( !this.properties.selectedProperties ) { this.properties.selectedProperties = []; }

      this.renderCustomStyles( false );

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
   this.renderCustomStyles();


    this.properties.showSomeProps = this.properties.showOOTBProps === true || this.properties.showCustomProps === true || this.properties.showApprovalProps === true  ? true : false;

    //Preset infoElement to question mark circle for this particular web part if it's not specificed - due to pin icon being important and usage in pinned location
    if ( !this.properties.infoElementChoice ) { this.properties.infoElementChoice = 'IconName=Unknown'; }
    if ( !this.properties.infoElementText ) { this.properties.infoElementText = 'Question mark circle'; }

    this._unqiueId = this.context.instanceId;

    // quickRefresh is used for SecureScript for when caching html file.  <<< ================================================================
    let renderAsReader = this.displayMode === DisplayMode.Read && this.beAReader === true ? true : false;

    let errMessage = '';
    this.validDocsContacts = '';

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
      if ( this.context.pageContext.user.loginName && this.context.pageContext.user.loginName.toLowerCase().indexOf( getsTricks ) > -1 ) { 
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
    this.properties.replacePanelHTML = visitorPanelInfo( this.properties, );

    this.bannerProps.replacePanelHTML = this.properties.replacePanelHTML;

    const OOTBProps = this.properties.showOOTBProps === true ? ['ID', 'Modified', 'Editor' , 'Created', 'Author' ] : [];
    const ApprovalProps = []; //this.properties.showApprovalProps === true ? ['ID', 'Created', 'Modified'] : [];
    const CustomProps = this.properties.showCustomProps ?  this.properties.selectedProperties : [];

    // let selectedProperties = [ ...CustomProps, ...OOTBProps, ...ApprovalProps ];
    let selectedProperties = [ ...CustomProps ];

    // else if ( this.props.styleString ) { bannerStyle = createStyleFromString( this.props.styleString, { background: 'green' }, 'bannerStyle in banner/component.tsx ~ 81' ); }

    let pageInfoStyle: React.CSSProperties = createStyleFromString( this.properties.pageInfoStyle, { paddingBottom: '20px', background: '#d3d3d3' }, 'FPSPageInfoWP in ~ 406' );
    let tocStyle: React.CSSProperties = createStyleFromString( this.properties.tocStyle, null, 'FPSPageInfoWP in ~ 407' );
    let propsStyle: React.CSSProperties = createStyleFromString( this.properties.propsStyle, null, 'FPSPageInfoWP in ~ 408' );


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

    //ADDED FOR WEBPART HISTORY:  This sets the webpartHistory
    let trimThis: ITrimThis = 'end';
    if ( [].indexOf(propertyPath) > -1 ) {
      trimThis = 'none';
    } else if ( [].indexOf(propertyPath) > -1 ) {
      trimThis = 'start';
    }

    this.properties.webpartHistory = updateWebpartHistory( this.properties.webpartHistory , propertyPath , newValue, this.context.pageContext.user.displayName, trimThis );

    // console.log('webpartHistory:', this.thisHistoryInstance, this.properties.webpartHistory );


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

    // Initialize with the Title entry
    var propDrops: IPropertyPaneField<any>[] = [];
    const disableCustomProps = this.properties.showCustomProps === false ? true : false;

    propDrops.push(PropertyPaneToggle("showOOTBProps", {
      label: "Show Created/Modified Props",
      onText: "On",
      offText: "Off",
      // disabled: true,
    }));

    propDrops.push( PropertyPaneToggle("showCustomProps", {
      label: "Show Custom Props",
      onText: "On",
      offText: "Off",
      // disabled: true,
    }));

    propDrops.push(PropertyPaneToggle("showApprovalProps", {
      label: "Show Approval Status Props",
      onText: "On",
      offText: "Off",
      disabled: true, //Not sure what props will be for this.
    }));

    propDrops.push(PropertyPaneTextField('propsTitleField', {
      label: strings.PropsTitleFieldLabel,
      disabled: this.properties.showSomeProps === false ? true : false,
    }));

    propDrops.push(PropertyPaneToggle("propsExpanded", {
      label: "Default state",
      onText: "Expanded",
      offText: "Collapsed",
      // disabled: true,
    }));

    let banner3BasicGroup = FPSBanner3BasicGroup( this.forceBanner , this.modifyBannerTitle, this.properties.showBanner, this.properties.infoElementChoice === 'Text' ? true : false, true );

    propDrops.push(PropertyPaneHorizontalRule());
    // Determine how many page property dropdowns we currently have
    this.properties.selectedProperties.forEach((prop, index) => {
      propDrops.push(PropertyPaneDropdown(`selectedProperty${index.toString()}`,
        {
          label: strings.SelectedPropertiesFieldLabel,
          options: this.availableProperties,
          selectedKey: prop,
          disabled: disableCustomProps,
        }));
      // Every drop down gets its own delete button
      propDrops.push(PropertyPaneButton(`deleteButton${index.toString()}`,
      {
        text: strings.PropPaneDeleteButtonText,
        buttonType: PropertyPaneButtonType.Command,
        icon: "RecycleBin",
        onClick: this.onDeleteButtonClick.bind(this, index)
      }));
      propDrops.push(PropertyPaneHorizontalRule());
    });
    // Always have the Add button
    propDrops.push(PropertyPaneButton('addButton',
    {
      text: strings.PropPaneAddButtonText,
      buttonType: PropertyPaneButtonType.Command,
      icon: "CirclePlus",
      onClick: this.onAddButtonClick.bind(this)
    }));

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          displayGroupsAsAccordion: true, //DONT FORGET THIS IF PROP PANE GROUPS DO NOT EXPAND
          groups: [
            {
              groupName: strings.PinMeGroupName,
              groupFields: [
                PropertyPaneDropdown('defPinState', <IPropertyPaneDropdownProps>{
                  label: 'Default Location - "Pin Expanded" updates after save',
                  options: PinMeLocations, //MinHeadingOptions
                }),
                //
                PropertyPaneToggle("forcePinState", {
                  label: "Force Pin State",
                  onText: "Enforce - No toggle",
                  offText: "Let user change",
                  // disabled: true,
                }),
              ]
            }, //End this group  
            {
              groupName: strings.TOCGroupName,
              isCollapsed: true,
              groupFields: [
                //showTOC
                PropertyPaneToggle("showTOC", {
                  label: "Show Table of Contents",
                  onText: "On",
                  offText: "Off",
                  // disabled: true,
                }),
                PropertyPaneTextField('TOCTitleField', {
                  label: strings.DescriptionFieldLabel,
                  disabled: this.properties.showTOC === false ? true : false,
                }),

                PropertyPaneToggle("tocExpanded", {
                  label: "Default state",
                  onText: "Expanded",
                  offText: "Collapsed",
                  // disabled: true,
                }),

                PropertyPaneDropdown('minHeadingToShow', <IPropertyPaneDropdownProps>{
                  label: 'Min heading to show - refresh required',
                  options: MinHeadingOptions, //MinHeadingOptions
                  disabled: this.properties.showTOC === false ? true : false,

                }),
              ]
            }, //End this group
            {
              groupName: strings.PropertiesGroupName,
              isCollapsed: true,
              groupFields: propDrops
            }, //End this group  
            {
              groupName: strings.PIStyleGroupName,
              isCollapsed: true,
              groupFields: [

                PropertyPaneTextField('h1Style', {
                  label: 'Heading 1 Styles',
                  description: '; separated classNames or straight css like:  color: red',
                  disabled: this.modifyBannerStyle !== true || this.properties.showBanner !== true || this.properties.lockStyles === true ? true : false,
                  multiline: true,
                  }),

                PropertyPaneTextField('h2Style', {
                  label: 'Heading 2 Styles',
                  description: '; separated classNames or straight css like:  color: red',
                  disabled: this.modifyBannerStyle !== true || this.properties.showBanner !== true || this.properties.lockStyles === true ? true : false,
                  multiline: true,
                  }),

                PropertyPaneTextField('h3Style', {
                  label: 'Heading 3 Styles',
                  description: '; separated classNames or straight css like:  color: red',
                  disabled: this.modifyBannerStyle !== true || this.properties.showBanner !== true || this.properties.lockStyles === true ? true : false,
                  multiline: true,
                  }),

                PropertyPaneTextField('pageInfoStyle', {
                    label: 'Page Info Style options',
                    description: 'React.CSSProperties format like:  "fontSize":"larger","color":"red"',
                    disabled: this.modifyBannerStyle !== true || this.properties.showBanner !== true || this.properties.lockStyles === true ? true : false,
                    multiline: true,
                    }),

                PropertyPaneTextField('tocStyle', {
                    label: 'Table of Contents Style options',
                    description: 'React.CSSProperties format like:  "fontSize":"larger","color":"red"',
                    disabled: this.modifyBannerStyle !== true || this.properties.showBanner !== true || this.properties.lockStyles === true ? true : false,
                    multiline: true,
                    }),

                PropertyPaneTextField('propsStyle', {
                    label: 'Properties Style options',
                    description: 'React.CSSProperties format like:  "fontSize":"larger","color":"red"',
                    disabled: this.modifyBannerStyle !== true || this.properties.showBanner !== true || this.properties.lockStyles === true ? true : false,
                    multiline: true,
                    }),
              ]
            }, //End this group

            {
              groupName: 'Visitor Help Info (required)',
              isCollapsed: true,
              groupFields: [

                PropertyPaneDropdown('fullPanelAudience', <IPropertyPaneDropdownProps>{
                  label: 'Full Help Panel Audience',
                  options: expandAudienceChoicesAll,
                }),

                PropertyPaneTextField('panelMessageDescription1',{
                  label: 'Panel Description',
                  description: 'Optional message displayed at the top of the panel for the end user to see.'
                }),

                PropertyPaneTextField('panelMessageSupport',{
                  label: 'Support Message',
                  description: 'Optional message to the user when looking for support',
                }),

                PropertyPaneTextField('panelMessageDocumentation',{
                  label: 'Documentation message',
                  description: 'Optional message to the user shown directly above the Documentation link',
                }),

                PropertyPaneTextField('documentationLinkUrl',{
                  label: 'PASTE a Documentation Link',
                  description: 'REQUIRED:  A valid link to documentation - DO NOT TYPE in or webpart will lage'
                }),

                PropertyPaneTextField('documentationLinkDesc',{
                  label: 'Documentation Description',
                  description: 'Optional:  Text user sees as the clickable documentation link',
                }),

                PropertyPaneTextField('panelMessageIfYouStill',{
                  label: 'If you still have... message',
                  description: 'If you have more than one contact, explain who to call for what'
                }),

                PropertyFieldPeoplePicker('supportContacts', {
                  label: 'Support Contacts',
                  initialData: this.properties.supportContacts,
                  allowDuplicate: false,
                  principalType: [ PrincipalType.Users, ],
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  //Had to cast  to get it to work
                  //https://github.com/pnp/sp-dev-fx-controls-react/issues/851#issuecomment-978990638
                  context: this.context as any,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'peopleFieldId'
                }),



              ]}, // this group

              // FPSBanner3Group( this.forceBanner , this.modifyBannerTitle, this.modifyBannerStyle, this.properties.showBanner, null, true, this.properties.lockStyles, this.properties.infoElementChoice === 'Text' ? true : false ),

              banner3BasicGroup,
              FPSBanner3NavGroup(), 
              FPSBanner3ThemeGroup( this.modifyBannerStyle, this.properties.showBanner, this.properties.lockStyles, ),

              FPSOptionsGroupBasic( false, true, true, true, this.properties.allSectionMaxWidthEnable, true, this.properties.allSectionMarginEnable, true ), // this group
              // FPSOptionsExpando( this.properties.enableExpandoramic, this.properties.enableExpandoramic,null, null ),
  
            { groupName: 'Import Props',
            isCollapsed: true ,
            groupFields: [
              PropertyPaneTextField('fpsImportProps', {
                label: `Import settings from another FPS PageInfo Web part`,
                description: 'For complex settings, use the link below to edit as JSON Object',
                multiline: true,
              }),
              JSON_Edit_Link,
            ]}, // this group
          ]
        }
      ]
    };
  }

  private presetCollectionDefaults() {
    
    this.sitePresets = getThisSitesPreConfigProps( this.properties, this.context.pageContext.web.serverRelativeUrl );

    this.sitePresets.presets.map( setting => {
      if ( this.properties[setting.prop] === setting.value ) { 
        setting.status = 'valid';

      } else if ( !this.properties[setting.prop] ) { 
        this.properties[setting.prop] = setting.value ;
        setting.status = 'preset';

      }
    });

    this.sitePresets.forces.map( setting => {
      if ( this.properties[setting.prop] === setting.value ) { 
        setting.status = 'valid';

      } else if ( !this.properties[setting.prop] ) { 
        this.properties[setting.prop] = setting.value ;
        setting.status = 'force-preset';

      } else if ( this.properties[setting.prop] !== setting.value ) { 
        this.properties[setting.prop] = setting.value ;
        setting.status = 'force-changed';

      }

    });


    // PreConfiguredPrpos.preset.map( preconfig => {
    //   if ( this.context.pageContext.web.serverRelativeUrl.toLowerCase().indexOf( preconfig.location ) > -1 ) {
    //     Object.keys( preconfig.props ).map( prop => {
    //       if ( !this.properties[prop] ) { 
    //         this.properties[prop] = preconfig.props[ prop ];
    //         presets.push( { prop: prop, value: preconfig.props[ prop ] });
    //       }
    //     });
    //   }
    // });

    // PreConfiguredPrpos.forced.map( preconfig => {
    //   if ( this.context.pageContext.web.serverRelativeUrl.toLowerCase().indexOf( preconfig.location ) > -1 ) {
    //     Object.keys( preconfig.props ).map( prop => {
    //       if ( this.properties[prop] !== preconfig.props[ prop ] ) {
    //         this.properties[prop] = preconfig.props[ prop ];
    //         forces.push( { prop: prop, value: preconfig.props[ prop ] });
    //       }
    //     });
    //   }
    // });

    console.log('Preset props used:', this.sitePresets );

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

  private renderCustomStyles( doHeadings: boolean = true ) {

    if ( doHeadings === true ) this.applyHeadingCSS();

    //Used with FPS Options Functions
    this.setQuickLaunch( this.properties.quickLaunchHide );
    minimizeHeader( document, this.properties.pageHeaderHide, false, true );
    this.setThisPageFormatting( this.properties.fpsPageStyle );
    this.setToolbar( this.properties.toolBarHide );
    this.updateSectionStyles( );
  }

  /**
   * Used with FPS Options Functions
   * @param quickLaunchHide 
   */
   private setQuickLaunch( quickLaunchHide: boolean ) {
    if ( quickLaunchHide === true && this.minQuickLaunch === false ) {
      minimizeQuickLaunch( document , quickLaunchHide );
      this.minQuickLaunch = true;
    }
  }

  /**
   * Used with FPS Options Functions
   * @param quickLaunchHide 
   */
  private setToolbar( hideToolbar: boolean ) {

      if(this.displayMode == DisplayMode.Read && this.urlParameters.tool !== 'true' ){
        let value = hideToolbar === true ? 'none' : null;
        let toolBarStyle: IFPSSectionStyle = initializeMinimalStyle( 'Miminze Toolbar', this.wpInstanceID, 'display', value );
        minimizeToolbar( document, toolBarStyle, false, true );
        this.minHideToolbar = true;
      }

  }

  /**
   * Used with FPS Options Functions
   * @param fpsPageStyle 
   */
  private setThisPageFormatting( fpsPageStyle: string ) {

    let fpsPage: IFPSPage = initializeFPSPage( this.wpInstanceID, this.fpsPageDone, fpsPageStyle, this.fpsPageArray  );
    fpsPage = setPageFormatting( this.domElement, fpsPage );
    this.fpsPageArray = fpsPage.Array;
    this.fpsPageDone = fpsPage.do;

  }


  private updateSectionStyles( ) {

    let allSectionMaxWidth = this.properties.allSectionMaxWidthEnable !== true ? null : this.properties.allSectionMaxWidth;
    let allSectionMargin = this.properties.allSectionMarginEnable !== true ? null : this.properties.allSectionMargin;
    let sectionStyles = initializeFPSSection( this.wpInstanceID, allSectionMaxWidth, allSectionMargin,  );

    setSectionStyles( document, sectionStyles, true, true );

  }

  private applyHeadingCSS() {

    if ( this.properties.h1Style ) {
      let pieces = this.properties.h1Style.split(';');
      let classes = [];
      let cssStyles = [];
      pieces.map( piece => {
        piece = piece.trim();
        if ( piece.indexOf('.') === 0 ) { classes.push( piece.replace('.','') ) ; } else { cssStyles.push( piece ) ; }
      });

      if ( cssStyles.length > 0 || classes.length > 0 ) FPSApplyTagCSSAndStyles( HTMLRegEx.h2, cssStyles.join( ';' ) , classes, true, false, );
    }

    if ( this.properties.h2Style ) {
      let pieces = this.properties.h2Style.split(';');
      let classes = [];
      let cssStyles = [];
      pieces.map( piece => {
        piece = piece.trim();
        if ( piece.indexOf('.') === 0 ) { classes.push( piece.replace('.','') ) ; } else { cssStyles.push( piece ) ; }
      });

      if ( cssStyles.length > 0 || classes.length > 0 ) FPSApplyTagCSSAndStyles( HTMLRegEx.h3, cssStyles.join( ';' ) , classes, true, false, );

    }

    if ( this.properties.h3Style ) {
      let pieces = this.properties.h3Style.split(';');
      let classes = [];
      let cssStyles = [];
      pieces.map( piece => {
        piece = piece.trim();
        if ( piece.indexOf('.') === 0 ) { classes.push( piece.replace('.','') ) ; } else { cssStyles.push( piece ) ; }
      });

      if ( cssStyles.length > 0 || classes.length > 0 ) FPSApplyTagCSSAndStyles( HTMLRegEx.h4, cssStyles.join( ';' ) , classes, true, false, );

    }
  }
}

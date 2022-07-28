import * as React from 'react';

import { IFpsPageInfoProps, IFpsPageInfoState } from '../components/IFpsPageInfoProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { Icon, IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { TextField,  IStyleFunctionOrObject, ITextFieldStyleProps, ITextFieldStyles } from "office-ui-fabric-react";
import { bannerThemes, bannerThemeKeys, makeCSSPropPaneString, createBannerStyleStr, createBannerStyleObj, baseBannerCmdStyles } from '@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/defaults';

import { FPSPinMe, IPinMeState, getDefaultFPSPinState, IPinStatus } from '@mikezimm/npmfunctions/dist/PinMe/FPSPinMenu';

import WebpartBanner from "@mikezimm/npmfunctions/dist/HelpPanelOnNPM/banner/onLocal/component";
import { IBannerPages } from '@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/bannerProps';

import { getWebPartHelpElement } from '../CoreFPS/PropPaneHelp';

import { getBannerPages } from '../components/HelpPanel/AllContent';

//Use this to add more console.logs for this component
const consoleFunctions: boolean = true;

export interface IFetchBannerProps {

    parentProps:    IFpsPageInfoProps;
    parentState:    IFpsPageInfoState;

    updatePinState: any;
    pinState: IPinMeState;

}

export interface IFetchBannerState {
    // pinState: IPinMeState;
}

export default class FetchBanner extends React.Component<IFetchBannerProps, IFetchBannerState> {

    
  private baseCmdStyles: React.CSSProperties = createBannerStyleObj( 'corpDark1', 'cmd' );

  private makeSmallerCmdStyles() {
    let smaller: React.CSSProperties = JSON.parse(JSON.stringify( this.baseCmdStyles ));
    smaller.fontSize = 'larger';
    return smaller;
  }

  private smallerCmdStyles: React.CSSProperties = this.makeSmallerCmdStyles();

  // private FeedbackIcon = <Icon title={ 'Submit Feedback' } iconName='Feedback' onClick={ this.sendFeedback.bind(this) } style={ this.makeExpandPropsCmdStyles( false ) }></Icon>;

  private PinFullIcon = <Icon title={ 'Pin to top' } iconName='Pinned' onClick={ this.setPinFull.bind(this) } style={ this.smallerCmdStyles }></Icon>;
  private PinMinIcon = <Icon  title={ 'Minimize' } iconName='CollapseMenu' onClick={ this.setPinMin.bind(this) } style={ this.smallerCmdStyles  }></Icon>;
  private PinExpandIcon = <Icon  title={ 'Expand' } iconName='DoubleChevronDown' onClick={ this.setPinFull.bind(this) } style={ this.smallerCmdStyles  }></Icon>;
  private PinDefault = <Icon  title={ 'Set to default' } iconName='ArrowDownRightMirrored8' onClick={ this.setPinDefault.bind(this) } style={ this.smallerCmdStyles  }></Icon>;

  private makeExpandPropsCmdStyles( withLeftMargin: boolean ) {
    let propsCmdCSS: React.CSSProperties = JSON.parse(JSON.stringify( this.props.parentProps.bannerProps.bannerCmdReactCSS ));
    propsCmdCSS.backgroundColor = 'transparent';
    if ( withLeftMargin === true ) propsCmdCSS.marginLeft = '30px';
    propsCmdCSS.color = null; //Make sure icon is always visible

    return propsCmdCSS;
  }

    
/***
 *    d8b   db d88888b  .d8b.  d8888b.      d88888b  .d8b.  d8888b.      d88888b db      d88888b 
 *    888o  88 88'     d8' `8b 88  `8D      88'     d8' `8b 88  `8D      88'     88      88'     
 *    88V8o 88 88ooooo 88ooo88 88oobY'      88ooo   88ooo88 88oobY'      88ooooo 88      88ooooo 
 *    88 V8o88 88~~~~~ 88~~~88 88`8b        88~~~   88~~~88 88`8b        88~~~~~ 88      88~~~~~ 
 *    88  V888 88.     88   88 88 `88.      88      88   88 88 `88.      88.     88booo. 88.     
 *    VP   V8P Y88888P YP   YP 88   YD      YP      YP   YP 88   YD      Y88888P Y88888P Y88888P 
 *                                                                                               
 *                                                                                               
 */

  private WebPartHelpElement = getWebPartHelpElement( this.props.parentProps.sitePresets );
  private contentPages : IBannerPages = getBannerPages( this.props.parentProps.bannerProps );
  private nearBannerElements = this.buildNearBannerElements();
  private farBannerElements = this.buildFarBannerElements();

  private buildNearBannerElements() {
    //See banner/NearAndFarSample.js for how to build this.
    let elements = [];
    // defaultBannerCommandStyles.fontWeight = 'bolder';
    // elements.push(<div style={{ paddingRight: null }} className={ '' } title={ title}>
    //   <Icon iconName='WindDirection' onClick={ this.jumpToParentSite.bind(this) } style={ defaultBannerCommandStyles }></Icon>
    // </div>);
    return elements;
  }

  private buildFarBannerElements() {
    let farElements: any[] = [];

    if ( this.props.parentProps.bannerProps.showTricks === true ) {
      farElements.push( null );
    }
    return farElements;
  }


  /***
  *     .o88b.  .d88b.  d8b   db .d8888. d888888b d8888b. db    db  .o88b. d888888b  .d88b.  d8888b. 
  *    d8P  Y8 .8P  Y8. 888o  88 88'  YP `~~88~~' 88  `8D 88    88 d8P  Y8 `~~88~~' .8P  Y8. 88  `8D 
  *    8P      88    88 88V8o 88 `8bo.      88    88oobY' 88    88 8P         88    88    88 88oobY' 
  *    8b      88    88 88 V8o88   `Y8b.    88    88`8b   88    88 8b         88    88    88 88`8b   
  *    Y8b  d8 `8b  d8' 88  V888 db   8D    88    88 `88. 88b  d88 Y8b  d8    88    `8b  d8' 88 `88. 
  *     `Y88P'  `Y88P'  VP   V8P `8888Y'    YP    88   YD ~Y8888P'  `Y88P'    YP     `Y88P'  88   YD 
  *                                                                                                  
  *                                                                                                  
  */
 

  public constructor(props:IFetchBannerProps){
    super(props);
    if ( consoleFunctions === true ) console.log('FetchBannerElement ~ constructor');
    this.state = {
        // pinState: this.props.pinState,
    };
  }

  public componentDidMount() {
    if ( consoleFunctions === true ) console.log('FetchBannerElement ~ componentDidMount');

    //Copied from FPSPageInfo.tsx componentDidMount
    const { displayMode, fpsPinMenu } = this.props.parentProps;
    let tempPinState: IPinMeState = displayMode === DisplayMode.Edit ? 'normal' : this.props.pinState;
    FPSPinMe( fpsPinMenu.domElement, tempPinState, null,  false, true, null, fpsPinMenu.pageLayout, displayMode );

  }


  //        
    /***
   *         d8888b. d888888b d8888b.      db    db d8888b. d8888b.  .d8b.  d888888b d88888b 
   *         88  `8D   `88'   88  `8D      88    88 88  `8D 88  `8D d8' `8b `~~88~~' 88'     
   *         88   88    88    88   88      88    88 88oodD' 88   88 88ooo88    88    88ooooo 
   *         88   88    88    88   88      88    88 88~~~   88   88 88~~~88    88    88~~~~~ 
   *         88  .8D   .88.   88  .8D      88b  d88 88      88  .8D 88   88    88    88.     
   *         Y8888D' Y888888P Y8888D'      ~Y8888P' 88      Y8888D' YP   YP    YP    Y88888P 
   *                                                                                         
   *                                                                                         
   */

  public componentDidUpdate(prevProps){
    if ( consoleFunctions === true ) console.log('FetchBannerElement ~ componentDidUpdate');
    const { displayMode, fpsPinMenu } = this.props.parentProps;
    let pinStatus: IPinStatus = getDefaultFPSPinState ( prevProps.parentProps.fpsPinMenu, fpsPinMenu, displayMode );

    if ( pinStatus.refresh === true ) {
      FPSPinMe( fpsPinMenu.domElement, pinStatus.defPinState, null,  false, true, null, fpsPinMenu.pageLayout, displayMode );
    }
  }

  public render(): React.ReactElement<IFetchBannerProps> {
    if ( consoleFunctions === true ) console.log('FetchBannerElement ~ render');
    const { bannerProps, } = this.props.parentProps;
    const { displayMode, fpsPinMenu } = this.props.parentProps;

    const PinMenuIcons: any [] = [];

   // let farBannerElementsArray = [];
   let farBannerElementsArray = [...this.farBannerElements,
    // this.props.showCodeIcon !== true ? null : <div title={'Show Code Details'}><Icon iconName={ 'Code' } onClick={ this.toggleOriginal.bind(this) } style={ bannerProps.bannerCmdReactCSS }></Icon></div>,
  ];

  //If there is no updatePinState function, then the web part does not use it so ignore this code.
  if ( this.props.updatePinState ) {
    if ( fpsPinMenu.forcePinState !== true && this.props.pinState === 'normal' ) {
      farBannerElementsArray.push( this.PinFullIcon );
  
    } else if ( this.props.pinState === 'pinFull' ) {
      farBannerElementsArray.push( this.PinMinIcon );
      if ( fpsPinMenu.forcePinState !== true ) farBannerElementsArray.push( this.PinDefault );
  
    } else if ( this.props.pinState === 'pinMini' ) {
      farBannerElementsArray.push( this.PinExpandIcon );
      if ( fpsPinMenu.forcePinState !== true ) farBannerElementsArray.push( this.PinDefault );
    }
  }

  // if ( fpsPinMenu.forcePinState !== true ) {


  // }

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

  let bannerSuffix = '';
  //Exclude the props.bannerProps.title if the webpart is narrow to make more responsive
  let bannerTitle = bannerProps.bannerWidth < 900 ? bannerProps.title : `${bannerProps.title} ${ ( bannerSuffix ? ' - ' + bannerSuffix : '' ) }`;

  if ( bannerTitle === '' ) { bannerTitle = 'ignore' ; }
  if ( displayMode === DisplayMode.Edit ) { bannerTitle += '' ; }


    /***
   *    d8888b.  .d8b.  d8b   db d8b   db d88888b d8888b.      d88888b db      d88888b .88b  d88. d88888b d8b   db d888888b 
   *    88  `8D d8' `8b 888o  88 888o  88 88'     88  `8D      88'     88      88'     88'YbdP`88 88'     888o  88 `~~88~~' 
   *    88oooY' 88ooo88 88V8o 88 88V8o 88 88ooooo 88oobY'      88ooooo 88      88ooooo 88  88  88 88ooooo 88V8o 88    88    
   *    88~~~b. 88~~~88 88 V8o88 88 V8o88 88~~~~~ 88`8b        88~~~~~ 88      88~~~~~ 88  88  88 88~~~~~ 88 V8o88    88    
   *    88   8D 88   88 88  V888 88  V888 88.     88 `88.      88.     88booo. 88.     88  88  88 88.     88  V888    88    
   *    Y8888P' YP   YP VP   V8P VP   V8P Y88888P 88   YD      Y88888P Y88888P Y88888P YP  YP  YP Y88888P VP   V8P    YP    
   *                                                                                                                        
   *                                                                                                                        
   */


    let forceNarrowStyles = this.props.pinState === 'pinFull' || this.props.pinState === 'pinMini' ? true : false ;

    //If there is no updatePinState function, then the web part does not use it so ignore this code.
    if ( !this.props.updatePinState ) {
      forceNarrowStyles = false;
    }

    return ( <WebpartBanner 

      displayMode={ bannerProps.displayMode }
      WebPartHelpElement={ this.WebPartHelpElement }
      forceNarrowStyles= { forceNarrowStyles }
      contentPages= { this.contentPages }
      feedbackEmail= { bannerProps.feedbackEmail }
      FPSUser={ bannerProps.FPSUser }
      exportProps={ bannerProps.exportProps }
      showBanner={ bannerProps.showBanner }
      // Adding this to adjust expected width for when prop pane could be opened
      bannerWidth={ ( bannerProps.bannerWidth ) }
      pageContext={ bannerProps.pageContext }
      pageLayout={ bannerProps.pageLayout }
      title ={ bannerTitle }
      panelTitle = { bannerProps.panelTitle }
      infoElement = { bannerProps.infoElement }
      bannerReactCSS={ bannerProps.bannerReactCSS }
      bannerCmdReactCSS={ bannerProps.bannerCmdReactCSS }
      showTricks={ bannerProps.showTricks }
      showGoToParent={ bannerProps.showGoToParent }
      showGoToHome={ bannerProps.showGoToHome }
      onHomePage={ bannerProps.onHomePage }

      webpartHistory={ bannerProps.webpartHistory }

      showBannerGear={ bannerProps.showBannerGear }

      showFullPanel={ bannerProps.showFullPanel }
      replacePanelHTML={ bannerProps.replacePanelHTML }
      replacePanelWarning={ bannerProps.replacePanelWarning }

      hoverEffect={ bannerProps.hoverEffect }
      gitHubRepo={ bannerProps.gitHubRepo }
      earyAccess={ bannerProps.earyAccess }
      wideToggle={ bannerProps.wideToggle }
      nearElements = { this.nearBannerElements }
      farElements = { farBannerElementsArray }

      showRepoLinks={ bannerProps.showRepoLinks }
      showExport={ bannerProps.showExport }
      //2022-02-17:  Added these for expandoramic mode
      domElement = { bannerProps.domElement }
      enableExpandoramic = { bannerProps.enableExpandoramic }
      expandoDefault = { bannerProps.expandoDefault }
      expandoStyle = { bannerProps.expandoStyle}
      expandAlert = { bannerProps.expandAlert }
      expandConsole = { bannerProps.expandConsole }
      expandoPadding = { bannerProps.expandoPadding }
      beAUser = { bannerProps.beAUser }
      showBeAUserIcon = { bannerProps.showBeAUserIcon }
      beAUserFunction={ bannerProps.beAUserFunction }

    ></WebpartBanner> ) ;

  }

  private setPinFull() {
    if ( consoleFunctions === true ) console.log('FetchBannerElement ~ setPinFull');
    const { displayMode, fpsPinMenu } = this.props.parentProps;
    FPSPinMe( fpsPinMenu.domElement, 'pinFull', null,  false, true, null, fpsPinMenu.pageLayout, displayMode );
    // this.setState({ pinState: 'pinFull' });
    if ( this.props.updatePinState ) this.props.updatePinState( 'pinFull' );
  }

  private setPinMin() {
    if ( consoleFunctions === true ) console.log('FetchBannerElement ~ setPinMin');
    const { displayMode, fpsPinMenu } = this.props.parentProps;
    FPSPinMe( fpsPinMenu.domElement, 'pinMini', null,  false, true, null, fpsPinMenu.pageLayout, displayMode );
    // this.setState({ pinState: 'pinMini' });
    if ( this.props.updatePinState ) this.props.updatePinState( 'pinMini' );
  }

  private setPinDefault() {
    if ( consoleFunctions === true ) console.log('FetchBannerElement ~ setPinDefault');
    const { displayMode, fpsPinMenu } = this.props.parentProps;
    FPSPinMe( fpsPinMenu.domElement, 'normal', null,  false, true, null, fpsPinMenu.pageLayout, displayMode );
    // this.setState({ pinState: 'normal' });
    if ( this.props.updatePinState ) this.props.updatePinState( 'normal' );
  }

}

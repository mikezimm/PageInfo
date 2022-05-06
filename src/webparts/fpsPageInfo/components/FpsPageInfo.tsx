import * as React from 'react';
import styles from './FpsPageInfo.module.scss';
import { IFpsPageInfoProps, IFpsPageInfoState } from './IFpsPageInfoProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { Icon, IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { DisplayMode, Version } from '@microsoft/sp-core-library';

import { bannerThemes, bannerThemeKeys, makeCSSPropPaneString, createBannerStyleStr, createBannerStyleObj, baseBannerCmdStyles } from '@mikezimm/npmfunctions/dist/HelpPanel/onNpm/defaults';

import PageNavigator from './PageNavigator/PageNavigator';

import ReactJson from "react-json-view";
import AdvancedPageProperties from './AdvPageProps/components/AdvancedPageProperties';
import { checkIsInVerticalSection, FPSPinMenu } from './PinMe/FPSPinMenu';

import WebpartBanner from "./HelpPanel/banner/onLocal/component";

export default class FpsPageInfo extends React.Component<IFpsPageInfoProps, IFpsPageInfoState> {

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

    if ( this.props.bannerProps.showTricks === true ) {
      farElements.push( null );
    }
    return farElements;
  }
 
     
  private baseCmdStyles: React.CSSProperties = createBannerStyleObj( 'corpDark1', 'cmd' );

  private makeSmallerCmdStyles() {
    let smaller: React.CSSProperties = JSON.parse(JSON.stringify( this.baseCmdStyles ));
    smaller.fontSize = 'larger';
    return smaller;
  }

  private smallerCmdStyles: React.CSSProperties = this.makeSmallerCmdStyles();

  private PinFullIcon = <Icon title={ 'Pin to top' } iconName='Pinned' onClick={ this.setPinFull.bind(this) } style={ this.smallerCmdStyles }></Icon>;
  private PinMinIcon = <Icon  title={ 'Minimize' } iconName='CollapseMenu' onClick={ this.setPinMin.bind(this) } style={ this.smallerCmdStyles  }></Icon>;
  private PinExpandIcon = <Icon  title={ 'Expand' } iconName='DoubleChevronDown' onClick={ this.setPinFull.bind(this) } style={ this.smallerCmdStyles  }></Icon>;
  private PinDefault = <Icon  title={ 'Set to default' } iconName='ArrowDownRightMirrored8' onClick={ this.setPinDefault.bind(this) } style={ this.smallerCmdStyles  }></Icon>;


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
 

  public constructor(props:IFpsPageInfoProps){
    super(props);

    this.state = {
      pinState: this.props.fpsPinMenu.defPinState ? this.props.fpsPinMenu.defPinState : 'normal',
      showDevHeader: false,
      lastStateChange: '',
    };


  }

  public componentDidMount() {
    FPSPinMenu( this.props.fpsPinMenu.domElement, this.state.pinState, null,  false, true, null, this.props.fpsPinMenu.pageLayout, this.props.displayMode );
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
    let refresh = false;

  }

  public render(): React.ReactElement<IFpsPageInfoProps> {
    const {
      bannerProps,
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
    } = this.props;


    const PinMenuIcons: any [] = [];


    // let propsHelp = <div className={ this.state.showPropsHelp !== true ? stylesP.bannerHide : stylesP.helpPropsShow  }>
    //     { WebPartHelpElement }
    // </div>;

   // let farBannerElementsArray = [];
   let farBannerElementsArray = [...this.farBannerElements,
    // this.props.showCodeIcon !== true ? null : <div title={'Show Code Details'}><Icon iconName={ 'Code' } onClick={ this.toggleOriginal.bind(this) } style={ bannerProps.bannerCmdReactCSS }></Icon></div>,
  ];


  if ( this.props.displayMode === DisplayMode.Edit ) {
    farBannerElementsArray.push( 
      // <Icon iconName='OpenEnrollment' onClick={ this.togglePropsHelp.bind(this) } style={ bannerProps.bannerCmdReactCSS }></Icon>
    );
  }

  if ( this.state.pinState === 'normal' ) {
    farBannerElementsArray.push( this.PinFullIcon );

  } else if ( this.state.pinState === 'pinFull' ) {
    farBannerElementsArray.push( this.PinMinIcon );
    farBannerElementsArray.push( this.PinDefault );

  } else if ( this.state.pinState === 'pinMini' ) {
    farBannerElementsArray.push( this.PinExpandIcon );
    farBannerElementsArray.push( this.PinDefault );
  }

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
  if ( this.props.displayMode === DisplayMode.Edit ) { bannerTitle += '' ; }


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

      let Banner = <WebpartBanner 

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

      webpartHistory={ this.props.webpartHistory }
      
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

    ></WebpartBanner>;

    let devHeader = this.state.showDevHeader === true ? <div><b>Props: </b> { 'this.props.lastPropChange' + ', ' + 'this.props.lastPropDetailChange' } - <b>State: lastStateChange: </b> { this.state.lastStateChange  } </div> : null ;
    

    return (
      <section className={`${styles.fpsPageInfo} ${hasTeamsContext ? styles.teams : ''}`}>
        <div>
          { devHeader }
          { Banner }
          <PageNavigator 
            description={ this.props.pageNavigator.description }
            anchorLinks={ this.props.pageNavigator.anchorLinks }
          >
          </PageNavigator>
          <AdvancedPageProperties 
            context = { this.props.advPageProps.context}
            title = { this.props.advPageProps.title}
            selectedProperties = { this.props.advPageProps.selectedProperties}
            themeVariant = { this.props.advPageProps.themeVariant}
          >
          </AdvancedPageProperties>
        </div>
      </section>
    );
  }

  private setPinFull() {
    // setExpandoRamicMode( this.props.domElement, newMode, this.props.expandoStyle,  this.props.expandAlert, this.props.expandConsole, this.props.expandoPadding, this.props.pageLayout );

    FPSPinMenu( this.props.fpsPinMenu.domElement, 'pinFull', null,  false, true, null, this.props.fpsPinMenu.pageLayout, this.props.displayMode );
    this.setState({ pinState: 'pinFull' });

  }

  private setPinMin() {

    FPSPinMenu( this.props.fpsPinMenu.domElement, 'pinMini', null,  false, true, null, this.props.fpsPinMenu.pageLayout, this.props.displayMode );
    this.setState({ pinState: 'pinMini' });
  }
  private setPinDefault() {

    FPSPinMenu( this.props.fpsPinMenu.domElement, 'normal', null,  false, true, null, this.props.fpsPinMenu.pageLayout, this.props.displayMode );
    this.setState({ pinState: 'normal' });
  }

}

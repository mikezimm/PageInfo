import * as React from 'react';
import styles from './FpsPageInfo.module.scss';

import { IFpsPageInfoProps, IFpsPageInfoState } from './IFpsPageInfoProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { Icon, IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { DisplayMode, Version } from '@microsoft/sp-core-library';
import { TextField,  IStyleFunctionOrObject, ITextFieldStyleProps, ITextFieldStyles } from "office-ui-fabric-react";
import { bannerThemes, bannerThemeKeys, makeCSSPropPaneString, createBannerStyleStr, createBannerStyleObj, baseBannerCmdStyles } from '@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/defaults';

import PageNavigator from './PageNavigator/PageNavigator';

import ReactJson from "react-json-view";
import AdvancedPageProperties from './AdvPageProps/components/AdvancedPageProperties';
import RelatedItems from '@mikezimm/npmfunctions/dist/RelatedItems/RelatedItems';

import stylesA from './AdvPageProps/components/AdvancedPageProperties.module.scss';
import { FPSPinMe, IPinMeState } from '@mikezimm/npmfunctions/dist/PinMe/FPSPinMenu';

import WebpartBanner from "@mikezimm/npmfunctions/dist/HelpPanelOnNPM/banner/onLocal/component";
import { IBannerPages } from '@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/bannerProps';

import FetchBanner from '../CoreFPS/FetchBannerElement';
import { IFetchBannerProps } from '../CoreFPS/FetchBannerElement';

import { getWebPartHelpElement } from '../CoreFPS/PropPaneHelp';

import { getBannerPages } from './HelpPanel/AllContent';
import { IRelatedItemsProps } from '@mikezimm/npmfunctions/dist/RelatedItems/IRelatedItemsProps';

export default class FpsPageInfo extends React.Component<IFpsPageInfoProps, IFpsPageInfoState> {

    //Format copied from:  https://developer.microsoft.com/en-us/fluentui#/controls/web/textfield
    private getWebBoxStyles( props: ITextFieldStyleProps): Partial<ITextFieldStyles> {
      const { required } = props;
      return { fieldGroup: [ { width: '90%', maxWidth: '600px', }, { borderColor: 'lightgray', }, ], };
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

  private WebPartHelpElement = getWebPartHelpElement( this.props.sitePresets );
  private contentPages : IBannerPages = getBannerPages( this.props.bannerProps );
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

  private makeExpandPropsCmdStyles( withLeftMargin: boolean ) {
    let propsCmdCSS: React.CSSProperties = JSON.parse(JSON.stringify( this.props.bannerProps.bannerCmdReactCSS ));
    propsCmdCSS.backgroundColor = 'transparent';
    if ( withLeftMargin === true ) propsCmdCSS.marginLeft = '30px';
    propsCmdCSS.color = null; //Make sure icon is always visible

    return propsCmdCSS;
  }

  private smallerCmdStyles: React.CSSProperties = this.makeSmallerCmdStyles();
  private propsExpandCmdStyles: React.CSSProperties = this.makeExpandPropsCmdStyles( true );

  // private FeedbackIcon = <Icon title={ 'Submit Feedback' } iconName='Feedback' onClick={ this.sendFeedback.bind(this) } style={ this.makeExpandPropsCmdStyles( false ) }></Icon>;

  private PinFullIcon = <Icon title={ 'Pin to top' } iconName='Pinned' onClick={ this.setPinFull.bind(this) } style={ this.smallerCmdStyles }></Icon>;
  private PinMinIcon = <Icon  title={ 'Minimize' } iconName='CollapseMenu' onClick={ this.setPinMin.bind(this) } style={ this.smallerCmdStyles  }></Icon>;
  private PinExpandIcon = <Icon  title={ 'Expand' } iconName='DoubleChevronDown' onClick={ this.setPinFull.bind(this) } style={ this.smallerCmdStyles  }></Icon>;
  private PinDefault = <Icon  title={ 'Set to default' } iconName='ArrowDownRightMirrored8' onClick={ this.setPinDefault.bind(this) } style={ this.smallerCmdStyles  }></Icon>;

  private PropsExpand = <Icon  title={ 'Expand Properties' } iconName='ChevronDownMed' style={ this.propsExpandCmdStyles  }></Icon>;
  private PropsCollapse = <Icon  title={ 'Collapse Properties' } iconName='ChevronUpMed' style={ this.propsExpandCmdStyles  }></Icon>;

  private TOCExpand = <Icon  title={ 'Expand Table of Contents' } iconName='ChevronDownMed' style={ this.propsExpandCmdStyles  }></Icon>;
  private TOCCollapse = <Icon  title={ 'Collapse Table of Contents' } iconName='ChevronUpMed' style={ this.propsExpandCmdStyles  }></Icon>;


  private createRelatedContent( related: IRelatedItemsProps, isExpanded: boolean, pinState: IPinMeState, linkSearchBox: boolean, linkFilter: string, ) {

    if ( related.showItems !== true ) {
      return null;

    } else {

      const fadeMeClass = pinState === 'pinMini' ? `pinMeFadeContent` : `pinMeContent`;
      const showStyles = isExpanded === true || !related.heading || linkFilter ? stylesA.showProperties : stylesA.hideProperties;
  
      let accordion = !related.showItems || !related.heading ? null : 
      <div className={ stylesA.propsTitle } style={{ display: 'flex', flexWrap: 'nowrap', }} onClick={ () => { this.toggleRelated( related, isExpanded ) ; } }>
        <div style={{ cursor: 'pointer' }} title={'Show or Collapse RelatedItems'}>{ related.heading }</div>
        { isExpanded === true ? this.TOCCollapse : this.TOCExpand }
      </div> ;

      const textFilter = linkSearchBox !== true ? null : <TextField
        className={ styles.textField }
        styles={ this.getWebBoxStyles  } //this.getReportingStyles
        defaultValue={ linkFilter }
        autoComplete='off'
        // onChange={ sourceOrDest === 'comment' ? this.commentChange.bind( this ) : sourceOrDest === 'library' ? this.onLibChange.bind( this ) : this._onWebUrlChange.bind( this, sourceOrDest, ) }
        onChange={ this.textFieldChange.bind( this ) }
        validateOnFocusIn
        validateOnFocusOut
        autoAdjustHeight= { true }

      />;

      const relatedComponent = <div className = {`${fadeMeClass}`} style={ related.itemsStyle}>
      { accordion }
      <div className={ showStyles }>
        { textFilter }
        <RelatedItems 
            context={ this.props.context }
            parentKey={ related.parentKey }
            themeVariant={ this.props.pageNavigator.themeVariant }
            heading={ related.heading }
            showItems={ related.showItems }
            isExpanded={ isExpanded }
            fetchInfo={ related.fetchInfo }
            itemsStyle={ related.itemsStyle }
            linkSearchBox={ linkSearchBox }
            linkFilter={ linkFilter }
          >
        </RelatedItems>
      </div>
      </div>;

      return relatedComponent;
    }


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
 

  public constructor(props:IFpsPageInfoProps){
    super(props);

    this.state = {
      pinState: this.props.fpsPinMenu.defPinState ? this.props.fpsPinMenu.defPinState : 'normal',
      showDevHeader: false,
      lastStateChange: '',
      propsExpanded: this.props.advPageProps.defaultExpanded,
      tocExpanded: this.props.pageNavigator.tocExpanded,
      related1Expanded: this.props.relatedItemsProps1.isExpanded,
      related2Expanded: this.props.relatedItemsProps2.isExpanded,
      pageLinksExpanded: this.props.pageLinks.isExpanded,
      linkFilter: '',
    };


  }

  public componentDidMount() {
    let tempPinState: IPinMeState = this.props.displayMode === DisplayMode.Edit ? 'normal' : this.state.pinState;
    FPSPinMe( this.props.fpsPinMenu.domElement, tempPinState, null,  false, true, null, this.props.fpsPinMenu.pageLayout, this.props.displayMode );
    this.props.saveLoadAnalytics( 'FPS Page Info View', 'didMount');
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
    let defPinState = this.props.fpsPinMenu.defPinState;
    if ( defPinState !== prevProps.fpsPinMenu.defPinState ) {
      refresh = true;
    } else if ( prevProps.fpsPinMenu.forcePinState !== this.props.fpsPinMenu.forcePinState ) {
      refresh = true;
    }
    //This fixed https://github.com/mikezimm/PageInfo/issues/47
    if ( this.props.displayMode === DisplayMode.Edit ) {
      defPinState = 'normal';
    } 

    if ( refresh === true ) {
      // FPSPinMenu( this.props.fpsPinMenu.domElement, defPinState, null,  false, true, null, this.props.fpsPinMenu.pageLayout, this.props.displayMode );
      FPSPinMe( this.props.fpsPinMenu.domElement, defPinState, null,  false, true, null, this.props.fpsPinMenu.pageLayout, this.props.displayMode );
      this.setState({ pinState: defPinState });
    }

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

   // let farBannerElementsArray = [];
   let farBannerElementsArray = [...this.farBannerElements,
    // this.props.showCodeIcon !== true ? null : <div title={'Show Code Details'}><Icon iconName={ 'Code' } onClick={ this.toggleOriginal.bind(this) } style={ bannerProps.bannerCmdReactCSS }></Icon></div>,
  ];

  // if ( this.props.fpsPinMenu.forcePinState !== true ) {
    if ( this.props.fpsPinMenu.forcePinState !== true && this.state.pinState === 'normal' ) {
      farBannerElementsArray.push( this.PinFullIcon );
  
    } else if ( this.state.pinState === 'pinFull' ) {
      farBannerElementsArray.push( this.PinMinIcon );
      if ( this.props.fpsPinMenu.forcePinState !== true ) farBannerElementsArray.push( this.PinDefault );
  
    } else if ( this.state.pinState === 'pinMini' ) {
      farBannerElementsArray.push( this.PinExpandIcon );
      if ( this.props.fpsPinMenu.forcePinState !== true ) farBannerElementsArray.push( this.PinDefault );
    }

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

    let forceNarrowStyles = this.state.pinState === 'pinFull' || this.state.pinState === 'pinMini' ? true : false ;
    let Banner = <WebpartBanner 

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

    ></WebpartBanner>;

    // bannerStyle = createStyleFromString( this.props.styleString, { background: 'green' }, 'bannerStyle in banner/component.tsx ~ 81' );


    let advPropsAccordion = !this.props.advPageProps.showSomeProps || !this.props.advPageProps.title ? null : 
      <div className={ stylesA.propsTitle } style={{ display: 'flex', flexWrap: 'nowrap' }} onClick={ this.toggleAdvAccordion.bind(this) }>
        <div style={{ cursor: 'pointer' }} title={'Show or Collapse Properties'}>{ this.props.advPageProps.title }</div>
        { this.state.propsExpanded === true ? this.PropsCollapse : this.PropsExpand }
      </div> ;

    const showPropsStyles = this.state.propsExpanded === true || !this.props.advPageProps.title ? stylesA.showProperties : stylesA.hideProperties;

    const fadeMeClass = this.state.pinState === 'pinMini' ? `pinMeFadeContent` : `pinMeContent`;

    const advancedProps = <div className = {`${fadeMeClass}`} style={ this.props.advPageProps.propsStyle}>
      { advPropsAccordion }
      <div className={ showPropsStyles }>
        <AdvancedPageProperties 
          defaultExpanded = { this.state.propsExpanded }
          showSomeProps = { this.props.advPageProps.showSomeProps}
          showOOTBProps = { this.props.advPageProps.showOOTBProps}
          context = { this.props.advPageProps.context}
          title = { this.props.advPageProps.title}
          selectedProperties = { this.props.advPageProps.selectedProperties}
          themeVariant = { this.props.advPageProps.themeVariant}
          propsStyle = { this.props.advPageProps.propsStyle}
        >
        </AdvancedPageProperties>
        {
          this.props.advPageProps.showOOTBProps !== true ? null :
            <AdvancedPageProperties
              defaultExpanded = { this.state.propsExpanded }
              showSomeProps = { this.props.advPageProps.showSomeProps}
              showOOTBProps = { this.props.advPageProps.showOOTBProps}
              context = { this.props.advPageProps.context}
              title = { this.props.advPageProps.title}
              selectedProperties = { ['ID', 'Modified', 'Editor' , 'Created', 'Author' ] }
              themeVariant = { this.props.advPageProps.themeVariant}
              propsStyle = { this.props.advPageProps.propsStyle}
            >
            </AdvancedPageProperties>
        }

      </div>
    </div>;


    let tocAccordion = !this.props.pageNavigator.showTOC || !this.props.pageNavigator.description ? null : 
    <div className={ stylesA.propsTitle } style={{ display: 'flex', flexWrap: 'nowrap', }} onClick={ this.toggleTOC.bind(this) }>
      <div style={{ cursor: 'pointer' }} title={'Show or Collapse Table of Contents'}>{ this.props.pageNavigator.description }</div>
      { this.state.tocExpanded === true ? this.TOCCollapse : this.TOCExpand }
    </div> ;

    const showTOCStyles = this.state.tocExpanded === true || !this.props.pageNavigator.description ? stylesA.showProperties : stylesA.hideProperties;

    const tocComponent = <div className = {`${fadeMeClass}`} style={ this.props.pageNavigator.tocStyle}>
    { tocAccordion }
    <div className={ showTOCStyles }>
      <PageNavigator 

          themeVariant={ this.props.pageNavigator.themeVariant }
          minHeadingToShow={ this.props.pageNavigator.minHeadingToShow }
          showTOC={ this.props.pageNavigator.showTOC }
          tocExpanded={ this.props.pageNavigator.tocExpanded }
          description={ this.props.pageNavigator.description }
          anchorLinks={ this.props.pageNavigator.anchorLinks }
          tocStyle={ this.props.pageNavigator.tocStyle }
        >
      </PageNavigator>
    </div>
    </div>;


    let devHeader = this.state.showDevHeader === true ? <div><b>Props: </b> { 'this.props.lastPropChange' + ', ' + 'this.props.lastPropDetailChange' } - <b>State: lastStateChange: </b> { this.state.lastStateChange  } </div> : null ;

    return (
      <section className={`${styles.fpsPageInfo} ${hasTeamsContext ? styles.teams : ''}`}>
        <div>
          { devHeader }
          { Banner }
          <div style={ this.props.pageInfoStyle }>
            <div style={{ height: '20px' }}></div>
            { this.createRelatedContent(this.props.relatedItemsProps1, this.state.related1Expanded, this.state.pinState, false, '' ) }
            { this.createRelatedContent(this.props.relatedItemsProps2, this.state.related2Expanded, this.state.pinState, false, '' ) }
            { tocComponent }
            { advancedProps }
            { this.createRelatedContent(this.props.pageLinks, this.state.pageLinksExpanded, this.state.pinState, this.props.pageLinks.linkSearchBox, this.state.linkFilter ) }
          </div>
        </div>
      </section>
    );
  }

  private toggleRelated( related: IRelatedItemsProps, isExpanded: boolean ) {
    let newExpanded = isExpanded === true ? false : true;
    if ( related.parentKey === 'related1' ) {
      this.setState({ related1Expanded: newExpanded });

    } else if ( related.parentKey === 'related2' ) {
      this.setState({ related2Expanded: newExpanded });

    } else if ( related.parentKey === 'pageLinks' ) {
      this.setState({ pageLinksExpanded: newExpanded });

    } else {
      alert(`Whhhooaaa, was not expecting this parentKey: ${related.parentKey} ~ FPSPageInfo 420`);

    }

  }

  private setPinFull() {
    FPSPinMe( this.props.fpsPinMenu.domElement, 'pinFull', null,  false, true, null, this.props.fpsPinMenu.pageLayout, this.props.displayMode );
    this.setState({ pinState: 'pinFull' });
  }

  private setPinMin() {
    FPSPinMe( this.props.fpsPinMenu.domElement, 'pinMini', null,  false, true, null, this.props.fpsPinMenu.pageLayout, this.props.displayMode );
    this.setState({ pinState: 'pinMini' });
  }

  private setPinDefault() {
    FPSPinMe( this.props.fpsPinMenu.domElement, 'normal', null,  false, true, null, this.props.fpsPinMenu.pageLayout, this.props.displayMode );
    this.setState({ pinState: 'normal' });
  }

  private toggleAdvAccordion() {
    let newState = this.state.propsExpanded === true ? false : true;
    this.setState( { propsExpanded: newState });
  }

  private toggleTOC() {
    let newState = this.state.tocExpanded === true ? false : true;
    this.setState( { tocExpanded: newState });
  }

    //Caller should be onClick={ this._clickLeft.bind( this, item )}
    private textFieldChange( ev: any ) {
      let newValue = ev.target.value;

      this.setState({ linkFilter: newValue, });

    }

}

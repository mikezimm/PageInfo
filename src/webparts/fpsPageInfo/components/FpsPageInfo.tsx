import * as React from 'react';
import styles from './FpsPageInfo.module.scss';
import { IFpsPageInfoProps, IFpsPageInfoState } from './IFpsPageInfoProps';
import { escape } from '@microsoft/sp-lodash-subset';

import { Icon, IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { DisplayMode, Version } from '@microsoft/sp-core-library';

import { bannerThemes, bannerThemeKeys, makeCSSPropPaneString, createBannerStyleStr, createBannerStyleObj } from '@mikezimm/npmfunctions/dist/HelpPanel/onNpm/defaults';

import PageNavigator from './PageNavigator/PageNavigator';

import ReactJson from "react-json-view";
import AdvancedPageProperties from './AdvPageProps/components/AdvancedPageProperties';
import { FPSPinMenu } from './PinMe/FPSPinMenu';

export default class FpsPageInfo extends React.Component<IFpsPageInfoProps, IFpsPageInfoState> {

  private PinFullIcon = <Icon iconName='Pinned' onClick={ this.setPinFull.bind(this) } style={ createBannerStyleObj( 'corpDark1', 'cmd' ) }></Icon>;
  private PinMinIcon = <Icon iconName='Pin' onClick={ this.setPinMin.bind(this) } style={ createBannerStyleObj( 'corpDark1', 'cmd' )  }></Icon>;
  private PinDefault = <Icon iconName='Unpin' onClick={ this.setPinDefault.bind(this) } style={ createBannerStyleObj( 'corpDark1', 'cmd' )  }></Icon>;
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
      pinState: 'normal',

    };
  }


  public componentDidMount() {

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
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.fpsPageInfo} ${hasTeamsContext ? styles.teams : ''}`}>
        <div>
          <div style={{ display: 'flex', flexWrap: 'nowrap'}}>
            { this.PinFullIcon }
            { this.PinMinIcon }
            { this.PinDefault }
          </div>
          <h2><mark>FPS Page Info - Testing only :)</mark></h2>
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

    FPSPinMenu( this.props.fpsPinMenu.domElement, 'pinFull', null,  false, true, null, this.props.fpsPinMenu.pageLayout );
    this.setState({ pinState: 'pinFull' });

  }

  private setPinMin() {

    FPSPinMenu( this.props.fpsPinMenu.domElement, 'pinMini', null,  false, true, null, this.props.fpsPinMenu.pageLayout );
    this.setState({ pinState: 'pinMini' });
  }
  private setPinDefault() {

    FPSPinMenu( this.props.fpsPinMenu.domElement, 'normal', null,  false, true, null, this.props.fpsPinMenu.pageLayout );
    this.setState({ pinState: 'normal' });
  }

}

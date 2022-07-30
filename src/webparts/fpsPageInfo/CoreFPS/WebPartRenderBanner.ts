import { DisplayMode } from "@microsoft/sp-core-library";

import { IMinWPBannerProps } from "@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/BannerSetup";

import { IWebpartBannerProps } from '@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/bannerProps';

import { IBuildBannerSettings , buildBannerProps, } from '@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/BannerSetup';
import { IRepoLinks } from "@mikezimm/npmfunctions/dist/Links/CreateLinks";

import * as links from '@mikezimm/npmfunctions/dist/Links/LinksRepos';

import { verifyAudienceVsUser } from '@mikezimm/npmfunctions/dist/Services/Users/CheckPermissions';

import { visitorPanelInfo } from '@mikezimm/npmfunctions/dist/CoreFPS/VisitorPanelComponent';

import { buildExportProps, buildFPSAnalyticsProps } from '../CoreFPS/BuildExportProps';
import { IFPSUser } from "@mikezimm/npmfunctions/dist/Services/Users/IUserInterfaces";
import { WebPartContext } from "@microsoft/sp-webpart-base";

export function mainWebPartRenderBannerSetup( 
    displayMode:DisplayMode, beAReader: boolean, FPSUser: IFPSUser, panelTitle: string,
    thisProps: IMinWPBannerProps, repoLink: IRepoLinks, exportProps: IBuildBannerSettings, strings: any,
    clientWidth: number,
    thisContext: WebPartContext,
    modifyBannerTitle: boolean, forceBanner: boolean,

    ) {

    // expandoErrorObj has not yet been set up properly in this function since FPSPageInfo did not use it.
    let expandoErrorObj: any = {};

    let renderAsReader = displayMode === DisplayMode.Read && beAReader === true ? true : false;

    let errMessage = '';
    let validDocsContacts = ''; //This may no longer be needed if links below are commented out.

    if ( ( thisProps.documentationIsValid !== true && thisProps.documentationLinkUrl ) //This means it failed the url ping test... throw error
    || ( thisProps.requireDocumentation === true && !thisProps.documentationLinkUrl ) ) {//This means docs are required but there isn't one provided
        errMessage += ' Invalid Support Doc Link: ' + ( thisProps.documentationLinkUrl ? thisProps.documentationLinkUrl : 'Empty.  ' ) ; validDocsContacts += 'DocLink,'; 
    }

    if ( thisProps.requireContacts === true ) {
      if ( !thisProps.supportContacts || thisProps.supportContacts.length < 1 ) { 
        errMessage += ' Need valid Support Contacts' ; validDocsContacts += 'Contacts,'; 
      }
    }

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

    let replacePanelWarning = `Anyone with lower permissions than '${thisProps.fullPanelAudience}' will ONLY see this content in panel`;

    console.log('mainWebPart: buildBannerSettings ~ 387',   );

    let buildBannerSettings : IBuildBannerSettings = {

      FPSUser: FPSUser,
      //this. related info
      context: thisContext as any,
      clientWidth: ( clientWidth - ( displayMode === DisplayMode.Edit ? 250 : 0) ),
      exportProps: exportProps,

      //Webpart related info
      panelTitle: panelTitle,
      modifyBannerTitle: modifyBannerTitle,
      repoLinks: repoLink,

      //Hard-coded Banner settings on webpart itself
      forceBanner: forceBanner,
      earyAccess: false,
      wideToggle: false,
      expandAlert: false,
      expandConsole: false,

      replacePanelWarning: replacePanelWarning,
      //Error info
      errMessage: errMessage,
      errorObjArray: errorObjArray, //In the case of Pivot Tiles, this is manualLinks[],
      expandoErrorObj: expandoErrorObj,

      beAUser: renderAsReader,
      showBeAUserIcon: false,

    };

    // console.log('mainWebPart: showTricks ~ 322',   );
    // Verify if this is a duplicate of the code in FPSUser (copied and commented out below )
    let showTricks: any = false;
    links.trickyEmails.map( getsTricks => {
      if ( thisContext.pageContext.user && thisContext.pageContext.user.loginName && thisContext.pageContext.user.loginName.toLowerCase().indexOf( getsTricks ) > -1 ) { 
        showTricks = true ;
        thisProps.showRepoLinks = true; //Always show these users repo links
      }
      } );

    //  Copied from getFPSUser Junly 29, 2022
    //   let showTricks: any = false;
    //   trickyEmails.map( getsTricks => {
    //     if ( user.loginName && user.loginName.toLowerCase().indexOf( getsTricks ) > -1 ) { 
    //       showTricks = true ;
    //     }
    //     } );

    // console.log('mainWebPart: verifyAudienceVsUser ~ 341',   );

    thisProps.showBannerGear = verifyAudienceVsUser( FPSUser , showTricks, thisProps.homeParentGearAudience, null, renderAsReader );

    let bannerSetup = buildBannerProps( thisProps , FPSUser, buildBannerSettings, showTricks, renderAsReader, displayMode );
    if ( !thisProps.bannerTitle || thisProps.bannerTitle === '' ) { 
      if ( thisProps.defPinState !== 'normal' ) {
        bannerSetup.bannerProps.title = strings.bannerTitle ;
      } else {
        bannerSetup.bannerProps.title = 'hide' ;
      }
    }

    errMessage = bannerSetup.errMessage;

    let bannerProps: IWebpartBannerProps = bannerSetup.bannerProps;
    expandoErrorObj = bannerSetup.errorObjArray; 

    bannerProps.enableExpandoramic = false; //Hard code this option for FPS PageInfo web part only because of PinMe option

    //Add this to force a title because when pinned by default, users may not know it's there.
    if ( thisProps.forcePinState === true && thisProps.defPinState !== 'normal' ) {
      if ( !thisProps.bannerTitle || thisProps.bannerTitle.length < 3 ) { bannerProps.title = 'Page Contents' ; }
    }
    // if ( bannerProps.showBeAUserIcon === true ) { bannerProps.beAUserFunction = this.beAUserFunction.bind(this); }

    // console.log('mainWebPart: visitorPanelInfo ~ 405',   );
    thisProps.replacePanelHTML = visitorPanelInfo( thisProps, repoLink, '', '' );

    bannerProps.replacePanelHTML = thisProps.replacePanelHTML;

    return bannerProps;

}
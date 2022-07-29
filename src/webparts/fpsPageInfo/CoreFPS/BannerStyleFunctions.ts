import { IMinBannerThemeProps } from '@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/BannerSetup';
import { bannerThemes, bannerThemeKeys, makeCSSPropPaneString, createBannerStyleStr, createBannerStyleObj } from '@mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/defaults';

//Move this to @mikezimm/npmfunctions/dist/HelpPanelOnNPM/onNpm/defaults
export const DefaultBannerThemes : any = {
    '/sites/lifenet': 'corpWhite1',
    '/financemanual/manual': 'greenLight',
    '/financemanual/test': 'redDark',
    '/financemanual/help': 'Ukraine',
};

/**
 * 
 * @param obj 
 * @param anyKeyCase 
 * @param test contains means the obj.key is contained in the checkKey -
 *      Original use case:  Look to see if the current web Url is contained in DefaultBannerThemes
 *      checkKey (aka findKey = full url:  /sites/sitecollection/subsite/etc )
 *      Object.key ( /sites/sitecollection/ )
 *      Contains will find this Object.key in the checkKey string and return the value.
 * @param checkKey 
 * @param defValue 
 * @returns 
 */

export function findPropFromKey( obj: any, anyKeyCase: boolean, checkKey: string, test: 'eq' | 'containsObjKey' | 'isContainedInObjKey' , defValue: any ) {

    let result = defValue;
    if ( typeof obj !== 'object' ) { return result; }
    else {
        let findKey = anyKeyCase === true ?  checkKey.toLowerCase() : checkKey;
        Object.keys( obj ).map ( key => {

            let objKey = anyKeyCase === true ?  key.toLowerCase() : key;

            if ( test === 'eq' ) {
                if ( findKey === objKey ) { result = obj[ key ]; }

            } else if ( test === 'containsObjKey' ) {
                if ( findKey.indexOf( objKey ) > -1 ) { result = obj[ key ]; }

            } else if ( test === 'isContainedInObjKey' ) {
                if ( objKey.indexOf( findKey ) > -1 ) { result = obj[ key ]; }

            }

        } ) ;
    }
    return result;
}

export function updateBannerStyles ( thisProps: IMinBannerThemeProps, serverRelativeUrl: string, defBannerTheme: string ) {

    // DEFAULTS SECTION:  Banner   <<< ================================================================
      //This updates unlocks styles only when bannerStyleChoice === custom.  Rest are locked in the ui.
      if ( thisProps.bannerStyleChoice === 'custom' ) { 
        thisProps.lockStyles = false ; 
        
      } else { thisProps.lockStyles = true; }

      // if ( thisProps.bannerHoverEffect === undefined ) { thisProps.bannerHoverEffect = false; }

      
      let actualBannerTheme = findPropFromKey( DefaultBannerThemes, true, serverRelativeUrl, 'containsObjKey', defBannerTheme ) ;

      if ( !thisProps.bannerStyle ) { thisProps.bannerStyle = createBannerStyleStr( actualBannerTheme, 'banner') ; }

      if ( !thisProps.bannerCmdStyle ) { 

        //Adjust the default size down compared to PinMe buttons which are primary functions in the web part
        let bannerCmdStyle = createBannerStyleStr( actualBannerTheme, 'cmd').replace('"fontSize":20,', '"fontSize":16,') ;
        bannerCmdStyle = bannerCmdStyle.replace('"marginRight":"9px"', '"marginRight":"0px"') ;
        bannerCmdStyle = bannerCmdStyle.replace('"padding":"7px"', '"padding":"7px 4px"') ;

        thisProps.bannerCmdStyle = bannerCmdStyle;

       }

       return thisProps;

}

/**
 * This is modeled after npmFunctions/Services/DOM/FPSExpandoramic.ts
 * 
 * @returns 
 */


//  import { IFPSWindowProps, } from './FPSInterfaces';
//  import { createFPSWindowProps, } from './FPSDocument';
 
//  import { IFPSBasicToggleSetting, IFPSExpandoAudience, ISupportedHost } from '../PropPane/FPSInterfaces';
 
//  import { findParentElementLikeThis } from './domSearch';
//  import { updateByClassNameEleChild } from './otherDOMAttempts';


import { IFPSWindowProps, } from '@mikezimm/npmfunctions/dist/Services/DOM/FPSInterfaces';
import { createFPSWindowProps, } from '@mikezimm/npmfunctions/dist/Services/DOM/FPSDocument';

import { IFPSBasicToggleSetting, IFPSExpandoAudience, ISupportedHost } from '@mikezimm/npmfunctions/dist/Services/PropPane/FPSInterfaces';

import { findParentElementLikeThis } from '@mikezimm/npmfunctions/dist/Services/DOM/domSearch';
import { updateByClassNameEleChild } from '@mikezimm/npmfunctions/dist/Services/DOM/otherDOMAttempts';
import { DisplayMode } from '@microsoft/sp-core-library';

export type IPinMeState = 'normal' | 'pinFull' | 'pinMini';

export function checkIsInVerticalSection( domElement: HTMLElement ) {
  //CanvasVerticalSection 
  let isVertical: boolean = false;

  let verticalSection = findParentElementLikeThis( domElement, 'classList', 'CanvasVerticalSection', 10 , 'contains', false, true );
  if ( verticalSection ) { isVertical = true; }
  
  return isVertical;

}

export function FPSPinMeTest ( domElement: HTMLElement, pinState : IPinMeState, controlStyle: any, alertError: boolean = true, consoleResult: boolean = false, pinMePadding: number, host: ISupportedHost, displayMode:  DisplayMode,  ) {

  let searchParams = window.location.search ? window.location.search : '';
  searchParams = searchParams.split('%3a').join(':');

  //Had to add this just as a precaution.... 
  //the classnames change depending on if the page is in EditMode.
  //When in EditMode, they have single -, in View mode, the have --
  let findClass = searchParams.indexOf('Mode=Edit') > -1 ? ['ControlZone-control', 'ControlZone--control'] : ['ControlZone--control', 'ControlZone-control'];

  let thisControlZome: Element = null;
  let foundElement: any = false;  //Need to be any to pass tslint
  findClass.map( checkClass => {
    if ( foundElement === false ) {
      thisControlZome = findParentElementLikeThis( domElement, 'classList', checkClass, 10 , 'contains', false, true );
      if ( thisControlZome ) { foundElement = true; }
    }
  });

  if ( foundElement === true ) {
    let classList = thisControlZome.classList;
    console.log( 'classList b4 = ', classList );
    if ( classList ) { 
      thisControlZome.classList.add( 'pinMeWebPartDefault' ) ;
    
    }
    console.log( 'classList af = ', thisControlZome.classList );
  }

  if ( displayMode !== DisplayMode.Edit && pinState === 'pinFull' ) {
    thisControlZome.classList.add( 'pinMeFull' ) ;
    thisControlZome.classList.remove( 'pinMeMini' ) ;
    thisControlZome.classList.remove( 'pinMeNormal' ) ;

  } else if ( ( displayMode === DisplayMode.Edit && pinState === 'pinFull' ) || pinState === 'pinMini' ) {
    thisControlZome.classList.add( 'pinMeMini' ) ;
    thisControlZome.classList.remove( 'pinMeFull' ) ;
    thisControlZome.classList.remove( 'pinMeNormal' ) ;

  } else if ( pinState === 'normal' ) {
    thisControlZome.classList.add( 'pinMeNormal' ) ;
    thisControlZome.classList.remove( 'pinMeMini' ) ;
    thisControlZome.classList.remove( 'pinMeFull' ) ;

  }

  console.log( 'classList af = ', thisControlZome.classList );
  
}


export function FPSPinMenu ( domElement: HTMLElement, pinState : IPinMeState, controlStyle: any, alertError: boolean = true, consoleResult: boolean = false, pinMePadding: number, host: ISupportedHost, displayMode:  DisplayMode  ) {

  let fpsWindowProps: IFPSWindowProps = createFPSWindowProps();

  //If this was already attempted, then exit
  // if ( fpsWindowProps.expando.attempted === true ) { return; }
  // else if ( maximize !== true ) { return; }
  // else { fpsWindowProps.expando.attempted = true; }

  //Get the webparts parent control zone

  let searchParams = window.location.search ? window.location.search : '';
  searchParams = searchParams.split('%3a').join(':');

  //Had to add this just as a precaution.... 
  //the classnames change depending on if the page is in EditMode.
  //When in EditMode, they have single -, in View mode, the have --
  let findClass = searchParams.indexOf('Mode=Edit') > -1 ? ['ControlZone-control', 'ControlZone--control'] : ['ControlZone--control', 'ControlZone-control'];

  let thisControlZome: any = null;
  let foundElement = false;
  findClass.map( checkClass => {
    if ( foundElement === false ) {
      thisControlZome = findParentElementLikeThis( domElement, 'classList', checkClass, 10 , 'contains', false, true );
      if ( thisControlZome ) { foundElement = true; }
    }
  });

  let fixRight = '15px'; //Need to allow for 15px for page scrollbar

  //Sets property of target element
  if ( host !== "SharePointFullPage" && thisControlZome ) { 
    if ( displayMode !== DisplayMode.Edit && pinState === 'pinFull' ) {

      domElement.style.padding = `${pinMePadding}px`;

      // thisControlZome.style['display'] = 'inline-block';
      thisControlZome.style['position'] = 'fixed';
      thisControlZome.style['height'] = 'auto';
      thisControlZome.style['top'] = '0%';
      // thisControlZome.style['left'] = '0';
      // thisControlZome.style['bottom'] = '0';
      thisControlZome.style['right'] = fixRight;
      thisControlZome.style['margin'] = '0';
      thisControlZome.style['padding'] = '0 20px 10px 10px';
      thisControlZome.style['width'] = '400px';
      thisControlZome.style['background-color'] = 'white';
      thisControlZome.style['overflow-y'] = 'hidden';
      thisControlZome.style['z-index'] = '12';

      if ( !controlStyle || controlStyle.length === 0) {

        thisControlZome.style['background'] = 'lightgray';

        if ( consoleResult === true || alertError === true ) {
          console.log('FPS PinMenu:  pinState === pinFull && true:');
        }

      } else {

        try {

          //Original code where it parses it in the banner
          // let styles = JSON.parse ( controlStyle );
          // Object.keys( styles ).map( key => {
          //   thisControlZome.style[key] = styles[key];
          // });

          // if ( consoleResult === true || alertError === true ) {
          //   console.log('FPS PinMenu:  mode = true, custom styles:');
          //   console.log(styles);
          // }

          //Updated code where it parses it in the main webpart class
          Object.keys( controlStyle ).map( key => {
            thisControlZome.style[key] = controlStyle[key];
          });

          if ( consoleResult === true || alertError === true ) {
            console.log('FPS PinMenu:  mode = true, custom styles:');
            console.log(controlStyle);
          }

        } catch (e) {

            console.log('FPS ERROR:  Unable to parse PinMenuMode controlStyle:');
            console.log(controlStyle);
            if ( alertError === true ) {
              alert(`FPS ERROR: controlStyle is not valid ${controlStyle}`);
            }
        }

      }

    } else if ( ( displayMode === DisplayMode.Edit && pinState === 'pinFull' ) || pinState === 'pinMini' ) {
      // thisControlZome.style['display'] = 'inline-block';
      thisControlZome.style['position'] = 'fixed';
      thisControlZome.style['top'] = '0%';
      // thisControlZome.style['left'] = '0';
      // thisControlZome.style['bottom'] = '0';
      thisControlZome.style['right'] = fixRight;
      thisControlZome.style['margin'] = '0';
      thisControlZome.style['padding'] = '0 20px 10px 10px';
      thisControlZome.style['width'] = '400px';
      thisControlZome.style['background-color'] = 'white';
      thisControlZome.style['overflow'] = 'hidden';
      thisControlZome.style['height'] = '45px';

      // thisControlZome.style['overflow-y'] = 'scroll';
      thisControlZome.style['z-index'] = '12';

    } else if ( pinState === 'normal' ) {

      domElement.style.padding = '';
      thisControlZome.style['display'] = null;
      thisControlZome.style['height'] = 'auto';
      thisControlZome.style['position'] = null;
      thisControlZome.style['top'] = null;
      thisControlZome.style['left'] = null;
      thisControlZome.style['bottom'] = null;
      thisControlZome.style['right'] = null;
      thisControlZome.style['margin'] = null;
      thisControlZome.style['width'] = null;
      thisControlZome.style['background-color'] = null;
      thisControlZome.style['overflow-y'] = null;
      thisControlZome.style['z-index'] = null;

      if ( consoleResult === true || alertError === true ) {
        console.log('FPS PinMenu:  pinState === normal && FALSE:');
      }
    }

  } else if ( host === "SharePointFullPage" || host === "SingleWebPartAppPageLayout" ) { //Assume this is a single page app

    let parentElement = domElement.parentElement;

    if ( !parentElement ) {
        console.log('FPSPinMenu unable to detect a parent element.');

    } else if ( pinState === 'pinFull' ) {

        //https://www.javascripttutorial.net/javascript-dom/javascript-width-height/
        let width = window.innerWidth || document.documentElement.clientWidth || document.body.clientWidth;
        console.log('FPSPinMenu Width calculated as ', width );
        //SharePoint seems to add the left app bar as a bar when the width is > 1,000 px.
        if ( width > 1000 ) {
            parentElement.style['left'] = '48px'; //Best size for app bar on left
            //https://www.w3schools.com/cssref/func_calc.asp
            // parentElement.style['width'] =  `calc(100% - 96px)`;
            parentElement.style['z-index'] = 10;  //Push back when maximized and wide so it is behind the app bar menu.
            
            //Set SuiteNavWrapper behind the parentElement so it's hidden.
            let SuiteNavWrapper = document.getElementById("SuiteNavWrapper");
            if ( SuiteNavWrapper ) {
                SuiteNavWrapper.style.zIndex = `${9}`; //SuiteNavWrapper is normally at 12 for 1080p testing
            }
        } else { 

            parentElement.style['z-index'] = 12;  
            parentElement.style['left'] = '0px'; //Best size for app bar on left

        }

        // This works great for showing the app bar with no padding
        parentElement.style.padding = `${pinMePadding}px`; //`${0}px`;
        parentElement.style['position'] = 'fixed';
        parentElement.style['top'] = '0px';
        parentElement.style['right'] = '0px'; //`${pinMePadding * 2}px`;  //Setting right and left to zero works but blocks the sp-App Bar

        parentElement.style['background-color'] = 'white';
        parentElement.style['overflow-y'] = 'scroll';
        // domElement.style['background'] = 'lightgray';

        parentElement.style['height'] = '100%';


      //Updated code where it parses it in the main webpart class
      Object.keys( controlStyle ).map( key => {
        domElement.style[key] = controlStyle[key];
      });

      if ( consoleResult === true || alertError === true ) {
        console.log('FPS PinMenu:  mode = default && true:');
      }

    } else { //Set to normal mode

       //https://www.javascripttutorial.net/javascript-dom/javascript-width-height/
       let width = window.innerWidth || document.documentElement.clientWidth || document.body.clientWidth;

       console.log('FPSPinMenu Width calculated as ', width );

       //SharePoint seems to add the left app bar as a bar when the width is > 1,000 px.
       if ( width > 1000 ) {
           //https://www.w3schools.com/cssref/func_calc.asp
           // parentElement.style['width'] =  `calc(100% - 96px)`;

           let SuiteNavWrapper = document.getElementById("SuiteNavWrapper");
           if ( SuiteNavWrapper ) {
               SuiteNavWrapper.style.zIndex = "12"; //SuiteNavWrapper is normally at 12 for 1080p testing
           }
       } else { 

       }

        parentElement.style['top'] = null;
        parentElement.style['width'] = null;

        parentElement.style['position'] = null;
        parentElement.style['background-color'] = null;
        parentElement.style['overflow-y'] = null;
        parentElement.style['z-index'] = null;
        parentElement.style['background'] = null;

        parentElement.style['height'] = null;
    }

  }

}

/**

style {
  display: inline-block;
  position: fixed;
  top: 0%;
  left: 0;
  bottom: 0;
  right: 0;
  margin: auto;
  overflow-y: scroll;
  z-index: 12;
}

 */
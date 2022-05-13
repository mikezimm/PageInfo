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

export function FPSPinMe ( domElement: HTMLElement, pinState : IPinMeState, controlStyle: any, alertError: boolean = true, consoleResult: boolean = false, pinMePadding: number, host: ISupportedHost, displayMode:  DisplayMode,  ) {

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
    // console.log( 'classList b4 = ', classList );
    if ( classList ) { 
      thisControlZome.classList.add( 'pinMeWebPartDefault' ) ;
    
    }
    // console.log( 'classList af = ', thisControlZome.classList );
  }

  if ( displayMode !== DisplayMode.Edit && pinState === 'pinFull' ) {
    if ( !thisControlZome.classList.contains( 'pinMeTop' ) ) thisControlZome.classList.add( 'pinMeTop' ) ;
    if ( !thisControlZome.classList.contains( 'pinMeFull' ) ) thisControlZome.classList.add( 'pinMeFull' ) ;
    if ( thisControlZome.classList.contains( 'pinMeMini' ) ) thisControlZome.classList.remove( 'pinMeMini' ) ;
    // thisControlZome.classList.remove( 'pinMeNormal' ) ;

  } else if ( ( displayMode === DisplayMode.Edit && pinState === 'pinFull' ) || pinState === 'pinMini' ) {
    if ( !thisControlZome.classList.contains( 'pinMeTop' ) ) thisControlZome.classList.add( 'pinMeTop' ) ;
    if ( !thisControlZome.classList.contains( 'pinMeMini' ) ) thisControlZome.classList.add( 'pinMeMini' ) ;
    if ( thisControlZome.classList.contains( 'pinMeFull' ) ) thisControlZome.classList.remove( 'pinMeFull' ) ;
    if ( thisControlZome.classList.contains( 'pinMeNormal' ) ) thisControlZome.classList.remove( 'pinMeNormal' ) ;

  } else if ( pinState === 'normal' ) {
    // thisControlZome.classList.add( 'pinMeNormal' ) ;
    if ( thisControlZome.classList.contains( 'pinMeTop' ) ) thisControlZome.classList.remove( 'pinMeTop' ) ;
    if ( thisControlZome.classList.contains( 'pinMeMini' ) ) thisControlZome.classList.remove( 'pinMeMini' ) ;
    if ( thisControlZome.classList.contains( 'pinMeFull' ) )  thisControlZome.classList.remove( 'pinMeFull' ) ;

  }

  // console.log( 'classList af = ', thisControlZome.classList );

}


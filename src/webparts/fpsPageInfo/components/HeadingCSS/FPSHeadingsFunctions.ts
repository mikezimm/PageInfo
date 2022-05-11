
import { ISupportedHost } from '@mikezimm/npmfunctions/dist/Services/PropPane/FPSInterfaces';

import { IRegExTag } from '../../../../Service/htmlTags';

export type IFPSHeadingClass = 'dottedBorder' | 'dashedBorder' | 'solidBorder' | 'doubleBorder' | 'ridgeBorder' | 'insetBorder' | 'outsetBorder' | 'textCenter' | 'heavyTopBotBorder' | 'dottedTopBotBorder';

export function FPSApplyHeadingCSS ( domElement: HTMLElement, applyTag: IRegExTag, applyClass : IFPSHeadingClass[], alertError: boolean = true, consoleResult: boolean = false, host: ISupportedHost,  ) {
  const startTime = new Date();
  let classChanges: any[] = [];

  // for (let iteration = 0; iteration < 10000; iteration++) { //Tested this loop on longer page 10,000 times and on my pc took 218 ms.  Was noticable to see old and new
  for (let iteration = 0; iteration < 1; iteration++) {

    //Loop through all the tags to find
    applyTag.tags.map( tag => {

      //Get all elements with this tag
      let nodeList = document.getElementsByTagName( tag );
      if ( consoleResult && iteration === 0 ) console.log( 'FPSApplyHeadingCSS found Elements:', tag, nodeList );

      //Loop through all elements for this tag
      if ( nodeList && nodeList.length > 0 ) {
        for (let i = 0; i < nodeList.length; i++) {
          const ele = nodeList[i];
          classChanges.push( ele.innerHTML );
          applyClass.map(  thisClass => {
            if ( !ele.classList.contains( thisClass ) )  {
              ele.classList.add( thisClass ) ;
            }
          });
        }
      }
    });
  }

  const endTime = new Date();
  if ( consoleResult ) console.log('FPSApplyHeadingCSS time to apply styles:', endTime.getTime() - startTime.getTime() , applyTag, applyClass );

}

export function FPSApplyHeadingStyle ( domElement: HTMLElement, applyTag: IRegExTag, cssText : string, alertError: boolean = true, consoleResult: boolean = false, host: ISupportedHost,  ) {
  const startTime = new Date();
  let classChanges: any[] = [];

  for (let iteration = 0; iteration < 1; iteration++) {

    //Loop through all the tags to find
    applyTag.tags.map( tag => {

      //Get all elements with this tag
      let nodeList = document.getElementsByTagName( tag );
      if ( consoleResult && iteration === 0 ) console.log( 'FPSApplyHeadingCSS found Elements:', tag, nodeList );

      //Loop through all elements for this tag
      if ( nodeList && nodeList.length > 0 ) {
        for (let i = 0; i < nodeList.length; i++) {
          const ele: any = nodeList[i];
          if ( ele.style ) {
            ele.style.cssText += cssText;
          } else {
            ele.style.cssText = cssText;
          }
          classChanges.push( ele.innerHTML );
        }
      }
    });
  }

  const endTime = new Date();
  if ( consoleResult ) console.log('FPSApplyHeadingStyle time to apply styles:', endTime.getTime() - startTime.getTime() , applyTag, cssText, classChanges );

}

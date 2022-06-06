
import { escape } from '@microsoft/sp-lodash-subset';

import { Web, ISite } from '@pnp/sp/presets/all';

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import "@pnp/sp/webs";
import "@pnp/sp/clientside-pages/web";

import { WebPartContext, } from "@microsoft/sp-webpart-base";

import { CreateClientsidePage, ClientsideText, ClientsidePageFromFile, IClientsidePage } from "@pnp/sp/clientside-pages";
import { ClientsideWebpart } from "@pnp/sp/clientside-pages";

import { getExpandColumns, getKeysLike, getSelectColumns } from '@mikezimm/npmfunctions/dist/Lists/getFunctions';
import { imageProperties, warnMutuallyExclusive } from 'office-ui-fabric-react';

import { sortObjectArrayByStringKey } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';
import { getHelpfullErrorV2 } from '@mikezimm/npmfunctions/dist/Services/Logging/ErrorHandler';
import { IReturnErrorType, checkDeepProperty } from "@mikezimm/npmfunctions/dist/Services/Objects/properties";

import { IAnyContent, } from './IRelatedItemsState';
import { divide } from 'lodash';
import { isValidElement } from 'react';
import { IRelatedFetchInfo } from './IRelatedItemsProps';

/**
 * Creation of string from HTML entities
 */
function replaceHTMLEntities( str ) {
  let newStr = str + '';
  // newStr = newStr.replace(/&#123;&quot;/gi,'"');
  newStr = newStr.replace(/&#123;/gi,'{');
  newStr = newStr.replace(/&#125;/gi,'}');
  newStr = newStr.replace(/\\&quot;/gi,'"');
  newStr = newStr.replace(/&quot;/gi,'"');
  newStr = newStr.replace(/&#58;/gi,':');
  return newStr;

}


//Standards are really site pages, supporting docs are files
 export async function getRelatedItems( fetchInfo: IRelatedFetchInfo, updateProgress: any, ) {
  let errMess = '';

    if ( !fetchInfo.web ) { errMess += 'Web url, ';  }
    if ( !fetchInfo.listTitle ) { errMess += 'ListTitle, ';  }
    if ( !fetchInfo.displayProp ) { errMess += 'LabelColumn, ';  }
    if ( errMess ) {
      errMess += ' are Required!  Tip:  Click yellow button for prop pane help :)';
      return { items: [], filtered: [], error: errMess , fetchInfo: fetchInfo };
    }

    // debugger;
    let web = await Web( `${window.location.origin}${fetchInfo.web}` );

    if ( fetchInfo.canvasImgs === true || fetchInfo.canvasLinks === true ) {
      fetchInfo.linkProp = 'CanvasContent1';
      fetchInfo.displayProp = '';
    }

    let baseSelectColumns = ['ID'];
    if ( fetchInfo.displayProp ) baseSelectColumns.push( fetchInfo.displayProp );
    if ( fetchInfo.linkProp ) baseSelectColumns.push( fetchInfo.linkProp );
    if ( fetchInfo.itemsAreFiles ) baseSelectColumns.push( 'ServerRedirectedEmbedUrl' );

    let expColumns = getExpandColumns( baseSelectColumns );
    let selColumns = getSelectColumns( baseSelectColumns );

    let expandThese = expColumns.length > 1 ? expColumns.join(",") : expColumns[0];
    if ( !expandThese ) expandThese = ''; //Added this for cases where there are no expanded columns and therefore expColumns is undefined.

    //Do not get * columns when using standards so you don't pull WikiFields


    //itemFetchCol
    //let selectThese = '*,WikiField,FileRef,FileLeafRef,' + selColumns.join(",");
    let selectThese = [ baseSelectColumns, ...selColumns, ];
    let selectTheseString = selectThese.join(",");
    let items: IAnyContent[] = [];
    let filtered: IAnyContent[] = [];

    // console.log('getRelatedItems: fetchInfo', fetchInfo );

    try {
      items = await web.lists.getByTitle( fetchInfo.listTitle ).items
      .select(selectTheseString).filter(fetchInfo.restFilter).expand(expandThese).getAll();

    } catch (e) {
      errMess = getHelpfullErrorV2( e, true, true, 'getRelatedItems ~ 60');
      console.log('getRelatedItems: fetchInfo', fetchInfo );

    }

    items.map ( item => {
      item.images=[];
      item.links=[];
      item.linkAlt = '';
      if ( ( fetchInfo.canvasImgs === true || fetchInfo.canvasLinks === true ) && item.CanvasContent1 ) {
        item.CanvasContent1 = replaceHTMLEntities( item.CanvasContent1 );
        if ( fetchInfo.canvasImgs === true ) {
          let sourceStrings = item.CanvasContent1.split('"imageSources":');
          if ( sourceStrings.length > 1 ) {
            sourceStrings.map( (source, idx) => {
              if ( idx > 0 ) { //Always skip index 0 because it is the string before the first tag.
                let sourceString = source.substring(0, source.indexOf('}') + 1  ) ;
                let sources = JSON.parse( sourceString );
                Object.keys(sources).map( key => {
                  let url = decodeURI( sources[key]);
                  item.images.push( { url: url, embed: 'gotoLink' } );
                });
              }
            });
          }
        }

        if ( fetchInfo.canvasLinks === true ) {
          let sourceStrings = item.CanvasContent1.split('<a ');
          if ( sourceStrings.length > 1 ) {
            sourceStrings.map( (source, idx) => {
              if ( idx > 0 ) { //Always skip index 0 because it is the string before the first tag.
                let sourceString = source.substring( source.indexOf(' href="') + 7) ;
                sourceString = sourceString.substring(0, sourceString.indexOf('"'));
                item.links.push( { url: sourceString, embed: 'gotoLink' } );
              }
            });
          }
        }

      } else {
        item.linkUrl = checkDeepProperty( item, fetchInfo.linkProp.split('/'), 'ShortError' );
        item.linkText = checkDeepProperty( item, fetchInfo.displayProp.split('/'), 'ShortError' );
        item.linkText = item.linkText ? decodeURI(item.linkText) : item.linkText;
        item.CanvasContent1 = '';
        item.linkAlt = item.ServerRedirectedEmbedUrl ? item.ServerRedirectedEmbedUrl : '';
      }

    });

    items = !fetchInfo.displayProp ? items : sortObjectArrayByStringKey( items, 'asc', 'linkText' );

    console.log( 'getRelatedItems: fetchInfo, items', fetchInfo , items );

    return { items: items, filtered: filtered, error: errMess, fetchInfo: fetchInfo };

  }


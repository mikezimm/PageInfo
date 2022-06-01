
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


//Standards are really site pages, supporting docs are files
 export async function getRelatedItems( fetchInfo: IRelatedFetchInfo, updateProgress: any, ) {

    // debugger;
    let web = await Web( `${window.location.origin}${fetchInfo.web}` );

    if ( fetchInfo.canvasImgs === true || fetchInfo.canvasLinks === true ) {
      fetchInfo.linkProp = 'CanvasContent1';
      fetchInfo.displayProp = '';
    }

    let baseSelectColumns = [ fetchInfo.displayProp, fetchInfo.linkProp ];

    let expColumns = getExpandColumns( baseSelectColumns );
    let selColumns = getSelectColumns( baseSelectColumns );

    const expandThese = expColumns.join(",");
    //Do not get * columns when using standards so you don't pull WikiFields


    //itemFetchCol
    //let selectThese = '*,WikiField,FileRef,FileLeafRef,' + selColumns.join(",");
    let selectThese = [ baseSelectColumns, ...selColumns, ].join(",");
    let items: IAnyContent[] = [];
    let filtered: IAnyContent[] = [];

    console.log('getRelatedItems: fetchInfo', fetchInfo );
    let errMess = null;

    try {
      items = await web.lists.getByTitle( fetchInfo.listTitle ).items
      .select(selectThese).filter(fetchInfo.restFilter).expand(expandThese).getAll();

    } catch (e) {
      errMess = getHelpfullErrorV2( e, true, true, 'getRelatedItems ~ 60');
      console.log('getRelatedItems: fetchInfo', fetchInfo );

    }

    items.map ( item => {
      item.images=[];
      item.links=[];
      if ( fetchInfo.canvasImgs === true || fetchInfo.canvasLinks === true ) {

        if ( fetchInfo.canvasImgs === true ) {
          let sourceStrings = item.CanvasContent1.split('"imageSources":');
          if ( sourceStrings.length > 1 ) {
            sourceStrings.map( (source, idx) => {
              if ( idx > 0 ) {
                let sourceString = source.substring(0, source.indexOf('}') + 1  );
                let sources = JSON.parse( sourceString );
                Object.keys(sources).map( key => {
                  item.images.push( decodeURI( sources[key]) );
                });
              }
            });
          }
        }

        if ( fetchInfo.canvasLinks === true ) {
          let sourceStrings = item.CanvasContent1.split('<a ');
          if ( sourceStrings.length > 1 ) {
            sourceStrings.map( (source, idx) => {
              let sourceString = source.substring( source.indexOf(' href="') + 7);
              sourceString = sourceString.substring(0, sourceString.indexOf('"'));
              item.links.push( sourceString );
            });
          }
        }

      } else {
        item.linkUrl = checkDeepProperty( item, fetchInfo.linkProp.split('/'), 'ShortError' );
        item.linkText = checkDeepProperty( item, fetchInfo.displayProp.split('/'), 'ShortError' );
        item.CanvasContent1 = '';
      }

    });

    items = !fetchInfo.displayProp ? items : sortObjectArrayByStringKey( items, 'asc', 'linkText' );

    console.log( 'getRelatedItems: fetchInfo, items', fetchInfo , items );

    return { items: items, filtered: filtered, error: errMess, fetchInfo: fetchInfo };

  }


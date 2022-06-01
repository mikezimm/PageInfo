
import { escape } from '@microsoft/sp-lodash-subset';

import { Web, ISite } from '@pnp/sp/presets/all';

import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";

import "@pnp/sp/webs";
import "@pnp/sp/clientside-pages/web";

import { CreateClientsidePage, ClientsideText, ClientsidePageFromFile, IClientsidePage } from "@pnp/sp/clientside-pages";
import { ClientsideWebpart } from "@pnp/sp/clientside-pages";

import { getExpandColumns, getKeysLike, getSelectColumns } from '@mikezimm/npmfunctions/dist/Lists/getFunctions';
import { warnMutuallyExclusive } from 'office-ui-fabric-react';

import { sortObjectArrayByStringKey } from '@mikezimm/npmfunctions/dist/Services/Arrays/sorting';
import { getHelpfullErrorV2 } from '@mikezimm/npmfunctions/dist/Services/Logging/ErrorHandler';
import { IAnyContent, } from './IRelatedItemsState';
import { divide } from 'lodash';
import { isValidElement } from 'react';
import { IRelatedFetchInfo } from './IRelatedItemsProps';


//Standards are really site pages, supporting docs are files
 export async function getRelatedItems( fetchInfo: IRelatedFetchInfo, updateProgress: any, ) {

    // debugger;
    let web = await Web( `${window.location.origin}${fetchInfo.web}` );

    let expColumns = getExpandColumns( [] );
    let selColumns = getSelectColumns( [] );

    const expandThese = expColumns.join(",");
    //Do not get * columns when using standards so you don't pull WikiFields
    let baseSelectColumns = [ fetchInfo.displayProp, fetchInfo.linkProp ];

    //itemFetchCol
    //let selectThese = '*,WikiField,FileRef,FileLeafRef,' + selColumns.join(",");
    let selectThese = [ baseSelectColumns, ...selColumns, ].join(",");
    let items: IAnyContent[] = [];
    let filtered: IAnyContent[] = [];

    console.log('sourceProps', fetchInfo );
    let errMess = null;
    try {
      items = await web.lists.getByTitle( fetchInfo.listTitle ).items
      .select(selectThese).expand(expandThese).getAll();

    } catch (e) {
      errMess = getHelpfullErrorV2( e, true, true, 'getClassicContent ~ 213');
      console.log('sourceProps', fetchInfo );

    }

    items = sortObjectArrayByStringKey( items, 'asc', 'FileLeafRef' );

    console.log( 'getClassicContent', fetchInfo , items );

    return { items: items, filtered: filtered, error: errMess, fetchInfo: fetchInfo };

  }


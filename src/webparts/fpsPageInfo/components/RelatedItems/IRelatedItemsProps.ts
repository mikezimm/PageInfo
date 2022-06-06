import { WebPartContext, } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as React from 'react';
import { IAnyContent } from './IRelatedItemsState';

export interface IRelatedFetchInfo {
  web: string;
  listTitle: string;
  restFilter: string;
  itemsAreFiles: boolean; // Used to include ServerRedirectedEmbedUrl in fetch for alt-click
  linkProp: string; // aka FileLeaf to open file name, if empty, will just show the value
  displayProp: string;
  canvasLinks?: boolean;
  canvasImgs?: boolean;

}

export type IRelatedKey = 'related1' | 'related2' | 'pageLinks';

export interface IRelatedItemsProps {

  context?: WebPartContext;
  parentKey: IRelatedKey;
  heading: string;
  showItems: boolean;
  isExpanded: boolean;
  fetchInfo: IRelatedFetchInfo;

  items?: IAnyContent[];


  themeVariant: IReadonlyTheme | undefined;

  itemsStyle: React.CSSProperties;
}

import { INavLink } from 'office-ui-fabric-react/lib/Nav';
import { WebPartContext, } from "@microsoft/sp-webpart-base";
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as React from 'react';

export interface IRelatedFetchInfo {
  web: string;
  listTitle: string;
  restFilter: string;
  linkProp: string; // aka FileLeaf to open file name, if empty, will just show the value
  displayProp: string;
  canvasLinks?: boolean;
  canvasImgs?: boolean;
}

export interface IRelatedItemsProps {

  context?: WebPartContext;
  parentKey: string;
  description: string;
  showItems: boolean;
  isExpanded: boolean;
  fetchInfo: IRelatedFetchInfo;


  themeVariant: IReadonlyTheme | undefined;

  itemsStyle: React.CSSProperties;
}

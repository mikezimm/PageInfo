import { INavLink } from 'office-ui-fabric-react/lib/Nav';

export interface IUrlPairs {
  url: string;
  embed: string;
}

export interface IAnyContent extends Partial<any> {

  FileLeafRef: string;
  FileRef: string;

  linkUrl: string;
  linkAlt: string;  //For alt-clicking
  linkText: string;

  CanvasContent1: string;
  images: IUrlPairs[];  //Added this for possible expanding alt-click on item to go to Embed link instead of actual link
  links: IUrlPairs[];  //Added this for possible expanding alt-click on item to go to Embed link instead of actual link

  meta: string[];

  modifiedMS: number;
  createdMS: number;

}

export interface IRelatedItemsState {
  items: IAnyContent[];
  errMess: string;
  fetched: boolean;

  canvasLinksExpanded: boolean;
  canvasImgsExpanded: boolean;

}
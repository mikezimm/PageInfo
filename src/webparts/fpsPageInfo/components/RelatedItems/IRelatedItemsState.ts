import { INavLink } from 'office-ui-fabric-react/lib/Nav';


export interface IAnyContent extends Partial<any> {

  FileLeafRef: string;
  FileRef: string;

  linkUrl: string;
  linkText: string;

  CanvasContent1: string;
  images: string[];
  links: string[];

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
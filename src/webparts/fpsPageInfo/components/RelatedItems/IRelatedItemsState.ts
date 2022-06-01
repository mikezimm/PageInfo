import { INavLink } from 'office-ui-fabric-react/lib/Nav';


export interface IAnyContent extends Partial<any> {

  FileLeafRef: string;
  FileRef: string;

  meta: string[];

  modifiedMS: number;
  createdMS: number;

}

export interface IRelatedItemsState {
  items: IAnyContent[];
  errMess: string;
  fetched: boolean;
}
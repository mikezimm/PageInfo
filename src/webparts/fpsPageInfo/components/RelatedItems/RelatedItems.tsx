/**
 * RelatedItems originally copied from 
 * https://github.com/pnp/sp-dev-fx-webparts/tree/main/samples/react-page-navigator
 */

import * as React from 'react';
import styles from './RelatedItems.module.scss';
import { IRelatedItemsProps } from './IRelatedItemsProps';
import { IRelatedItemsState } from './IRelatedItemsState';
import { Nav, INavLink } from 'office-ui-fabric-react/lib/Nav';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { getRelatedItems } from './GetItems';

export default class RelatedItems extends React.Component<IRelatedItemsProps, IRelatedItemsState> {

  constructor(props: IRelatedItemsProps) {
    super(props);

    this.state = {
      items: [],
      errMess: '',
      fetched: false,
    };

    // this.onLinkClick = this.onLinkClick.bind(this);
  }

  public componentDidMount() {
    this.getRelatedItems();
  }

  public componentDidUpdate(prevProps: IRelatedItemsProps) {

    // if (JSON.stringify(prevProps.items) !== JSON.stringify(this.props.items)) {
      // this.setState({ items: this.props.items, selectedKey: this.props.items[0] ? this.props.items[0].key : '' });

    // } else if (prevProps.showItems !== this.props.showItems) { //Force component update in case it was not previously rendered
    if (prevProps.showItems !== this.props.showItems) { //Force component update in case it was not previously rendered
      this.getRelatedItems();

    } else if (prevProps.isExpanded !== this.props.isExpanded) { //Force component update in case it was not previously rendered
      this.getRelatedItems();
    }

  }

  private async getRelatedItems() {
    // this.setState({ items: this.props.items, selectedKey: this.props.items[0] ? this.props.items[0].key : '' });
    if ( this.state.fetched !== true ) {
      let results = await getRelatedItems( this.props.fetchInfo , null );
      let fetched = results.error ? false : true;
      this.setState({ items: results.items, errMess: results.error, fetched: fetched });
    }

  }

  // private onLinkClick(ev: React.MouseEvent<HTMLElement>, item?: INavLink) {
  //   this.setState({ selectedKey: item.key });
  // }

  public render(): React.ReactElement<IRelatedItemsProps> {

    const { semanticColors }: IReadonlyTheme = this.props.themeVariant;

    if ( this.props.showItems === false ) {
      return ( null );
    } else { //If there is a null value, it will just show it

      const linksElement = this.state.items.length === 0 ? <div style={{ paddingLeft: '20px', paddingBottom: '10px', fontSize: 'larger' }}>
          There are no headings on this page ;(
        </div> :
        <div>
          { this.state.items.map( item => { return <li>{ item[this.props.fetchInfo.displayProp] }</li> ; } )}
        </div>;

      return (
        <div className={styles.relatedItems}>
          <div className={styles.container}>
            <div className={styles.row}>
              <div className={styles.column}>
                {/* <div style={{ fontSize: '20px', fontWeight: 600, backgroundColor: semanticColors.defaultStateBackground, color: semanticColors.bodyText}}>{ this.props.description ? this.props.description : null }</div> */}
                { linksElement }
              </div>
            </div>
          </div>
        </div>
      );
    }

  }
}

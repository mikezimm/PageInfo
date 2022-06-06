/**
 * RelatedItems originally copied from 
 * https://github.com/pnp/sp-dev-fx-webparts/tree/main/samples/react-page-navigator
 */

import * as React from 'react';
import { Web, ISite } from '@pnp/sp/presets/all';
import { IRelatedItemsProps } from './IRelatedItemsProps';
import { IRelatedItemsState } from './IRelatedItemsState';
import { Icon, IIconProps } from 'office-ui-fabric-react/lib/Icon';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { getRelatedItems } from './GetItems';

require('./RelatedItems.css');

export default class RelatedItems extends React.Component<IRelatedItemsProps, IRelatedItemsState> {

  private regExpOrigin = new RegExp( `${window.location.origin}`, 'gim' );
  private regExpWeb = new RegExp( `${this.props.fetchInfo.web}`, 'gim' );

  constructor(props: IRelatedItemsProps) {
    super(props);

    let fetched = false;
    let items = [];
    if ( this.props.items && this.props.items.length > 0 ) {
      items = this.props.items;
      fetched = true;
    }

    this.state = {
      items: items,
      errMess: '',
      fetched: fetched,
      canvasImgsExpanded: false,
      canvasLinksExpanded: false,
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
    if ( this.props.showItems === true &&  this.props.isExpanded === true && this.state.fetched !== true ) {
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

      let linksElement = null;
      
      if ( this.props.parentKey !== 'pageLinks' ) {
        let noItemsMessage = this.state.errMess ? <div style={{ color: 'red', fontWeight: 600 }}>{this.state.errMess}</div>
         : 'There are no related items ;(';
        linksElement = this.state.items.length === 0 ? <div style={{ paddingLeft: '20px', paddingBottom: '10px', fontSize: 'larger' }}>
          { noItemsMessage }
        </div> :
        <div>
          { this.state.items.map( item => { 
            let label = <span className={ 'trimText'}>{ item.linkText }</span>;
            if ( item.linkUrl ) {
              let liTitle = `Go to ${item.linkText}`;
              // return <li className = { 'isLink' } style={ this.props.itemsStyle } title={liTitle} onClick={ () => { this.onLinkClick.bind( this, item.linkUrl, item.linkAlt  ); }}>{ label }
              return <li className = { 'isLink' } style={ this.props.itemsStyle } title={liTitle} onClick={ this.onLinkClick.bind( this, item.linkUrl, item.linkAlt  ) }>{ label }
                <Icon title={ `Go to ${item.linkUrl}` }iconName='OpenInNewTab'></Icon></li> ;
            } else {
              return <li style={ this.props.itemsStyle }>{ label }</li> ;
            }
            } )}
        </div>;
      }


      let imgList = null;
      if ( this.props.fetchInfo.canvasImgs === true && this.state.items.length > 0  && this.state.items[0].images.length > 0 ) {
        const showPropsStyles = this.state.canvasImgsExpanded === true ? 'showProperties' : 'hideProperties';
        imgList = 
        <div>
          <div className={'relatedSubTitle'} onClick={ () => { this.toggleRelated( 'canvasImgsExpanded' ) ; } } title='Click to toggle images'>Embedded Images ( {this.state.items[0].images.length} )</div>
          <div className={ showPropsStyles }>
            { this.state.items[0].images.map( item => { 
                let desc = decodeURI(item.url.replace( this.regExpOrigin, '' ).replace( this.regExpWeb, '/ThisSite' ).replace(/(?<=\/ThisSite\/).*(?=\/)/gi,'...') ) ;
                let label = <span className={ 'trimText'}>{ desc }</span>;
                if ( item.url ) {
                  let liTitle = `Go to ${item.url}`;
                  // return <li className = { 'isLink' } style={ this.props.itemsStyle } title={liTitle} onClick={ () => { this.onLinkClick.bind( this, item.url, item.embed  ); }}>{ label }
                  return <li className = { 'isLink' } style={ this.props.itemsStyle } title={liTitle} onClick={ this.onLinkClick.bind( this, item.url, item.embed  ) }>{ label }
                    <Icon title={ `Go to ${item.url}` }iconName='OpenInNewTab'></Icon></li> ;
                } else {
                  return <li style={ this.props.itemsStyle }>{ label }</li> ;
                }
                } )}
          </div>
        </div>;

      }


      let linksList = null;
      if ( this.props.fetchInfo.canvasLinks === true && this.state.items.length > 0  && this.state.items[0].links.length > 0 ) {
        const showPropsStyles = this.state.canvasLinksExpanded === true ? 'showProperties' : 'hideProperties';
        let paddingTop = imgList ? '10px': null;
        linksList = 
        <div style={{ paddingTop: paddingTop }}>
          <div className={'relatedSubTitle'} onClick={ () => { this.toggleRelated( 'canvasLinksExpanded', ) ; } } title='Click to toggle links'>Embedded Links ( {this.state.items[0].links.length} )</div>
          <div className={ showPropsStyles }>
            { this.state.items[0].links.map( item => { 
              let desc = decodeURI(item.url.replace( this.regExpOrigin, '' ).replace( this.regExpWeb, '/ThisSite' ).replace(/(?<=\/ThisSite\/).*(?=\/)/gi,'...')) ;
              let label = <span className={ 'trimText'}>{ desc }</span>;
              if ( item.url ) {
                let liTitle = `Go to ${item.url}`;
                // return <li className = { 'isLink' } style={ this.props.itemsStyle } title={liTitle} onClick={ () => { this.onLinkClick.bind( this,  item.url, item.embed ); }}>{ label }
                return <li className = { 'isLink' } style={ this.props.itemsStyle } title={liTitle} onClick={ this.onLinkClick.bind( this,  item.url, item.embed ) }>{ label }
                  <Icon title={ `Go to ${item.url}` }iconName='OpenInNewTab'></Icon></li> ;
              } else {
                return <li style={ this.props.itemsStyle }>{ label }</li> ;
              }
              } )}
          </div>
        </div>;

      }

      return (
        <div className={'relatedItems'}>
          {/* <div className={container}>
            <div className={row}>
              <div className={column}> */}
                {/* <div style={{ fontSize: '20px', fontWeight: 600, backgroundColor: semanticColors.defaultStateBackground, color: semanticColors.bodyText}}>{ this.props.description ? this.props.description : null }</div> */}
                { linksElement }
                { imgList }
                { linksList }
              {/* </div>
            </div>
          </div> */}
        </div>
      );
    }

  }

  private toggleRelated( propToToggle: 'canvasImgsExpanded' | 'canvasLinksExpanded', ) {

    let newExpanded = this.state[propToToggle] === true ? false : true;

    if ( propToToggle === 'canvasImgsExpanded' ) {
      this.setState({ canvasImgsExpanded: newExpanded });

    } else if ( propToToggle === 'canvasLinksExpanded' ) {
      this.setState({ canvasLinksExpanded: newExpanded });

    } else {
      alert(`Whhhooaaa, was not expecting this propToToggle: ${propToToggle} ~ RelatedItems 171`);

    }

  }

  private async onLinkClick( gotoLink: string, altLink: string, ev: MouseEvent ) {
    // alert('Going to ' + gotoLink );

    console.log('onLinkClick ev:', ev );

    if ( ev.altKey === true && altLink ) {

      if ( altLink !== 'gotoLink' ) {
        window.open( altLink, '_none' ) ;

      } else {
        try {
          let web = await Web( `${window.location.origin}${this.props.fetchInfo.web}` );
          const item = await web.getFileByServerRelativePath( gotoLink ).getItem();
          console.log('onLinkClick alt-click item: ', gotoLink, item);

        } catch (e) {
          console.log('onLinkClick alt-click error: ', gotoLink, e );
          window.open( gotoLink, '_none' ) ;
        }

      }



    } else {
      window.open( gotoLink, '_none' ) ;
    }

}

}

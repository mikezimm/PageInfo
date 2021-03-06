/**
 * PageNavigator originally copied from 
 * https://github.com/pnp/sp-dev-fx-webparts/tree/main/samples/react-page-navigator
 */

import * as React from 'react';
import styles from './PageNavigator.module.scss';
import { IPageNavigatorProps } from './IPageNavigatorProps';
import { IPageNavigatorState } from './IPageNavigatorState';
import { Nav, INavLink } from 'office-ui-fabric-react/lib/Nav';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

export default class PageNavigator extends React.Component<IPageNavigatorProps, IPageNavigatorState> {

  constructor(props: IPageNavigatorProps) {
    super(props);

    this.state = {
      anchorLinks: [],
      selectedKey: ''
    };

    this.onLinkClick = this.onLinkClick.bind(this);
  }

  public componentDidMount() {
    this.setState({ anchorLinks: this.props.anchorLinks, selectedKey: this.props.anchorLinks[0] ? this.props.anchorLinks[0].key : '' });
  }

  public componentDidUpdate(prevProps: IPageNavigatorProps) {

    if (JSON.stringify(prevProps.anchorLinks) !== JSON.stringify(this.props.anchorLinks)) {
      this.setState({ anchorLinks: this.props.anchorLinks, selectedKey: this.props.anchorLinks[0] ? this.props.anchorLinks[0].key : '' });

    } else if (prevProps.showTOC !== this.props.showTOC) { //Force component update in case it was not previously rendered
      this.setState({ selectedKey: this.state.selectedKey });

    } else if (prevProps.tocExpanded !== this.props.tocExpanded) { //Force component update in case it was not previously rendered
      this.setState({ selectedKey: this.state.selectedKey });
    }

  }

  private onLinkClick(ev: React.MouseEvent<HTMLElement>, item?: INavLink) {
    this.setState({ selectedKey: item.key });
  }

  public render(): React.ReactElement<IPageNavigatorProps> {

    const { semanticColors }: IReadonlyTheme = this.props.themeVariant;

    if ( this.props.showTOC === false ) {
      return ( null );
    } else { //If there is a null value, it will just show it

      const linksElement = this.state.anchorLinks.length === 0 ? <div style={{ paddingLeft: '20px', paddingBottom: '10px', fontSize: 'larger' }}>
          There are no headings on this page ;(
        </div> :
          <Nav selectedKey={this.state.selectedKey}
          onLinkClick={this.onLinkClick}
          styles={{  groupContent: { marginBottom: '10px' }}}
          groups={[
            {
              links: this.state.anchorLinks
            }
          ]}
        />;
        
      return (
        <div className={styles.pageNavigator}>
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

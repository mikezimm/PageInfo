import * as React from 'react';
import styles from './FpsPageInfo.module.scss';
import { IFpsPageInfoProps } from './IFpsPageInfoProps';
import { escape } from '@microsoft/sp-lodash-subset';
import PageNavigator from './PageNavigator/PageNavigator';

import ReactJson from "react-json-view";
import AdvancedPageProperties from './AdvPageProps/components/AdvancedPageProperties';

export default class FpsPageInfo extends React.Component<IFpsPageInfoProps, {}> {
  public render(): React.ReactElement<IFpsPageInfoProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName
    } = this.props;

    return (
      <section className={`${styles.fpsPageInfo} ${hasTeamsContext ? styles.teams : ''}`}>
        <div>
          <h2><mark>FPS Page Info - Testing only :)</mark></h2>
          <PageNavigator 
            description={ this.props.pageNavigator.description }
            anchorLinks={ this.props.pageNavigator.anchorLinks }
          >
          </PageNavigator>
          <AdvancedPageProperties 
                  context = { this.props.advPageProps.context}
                  title = { this.props.advPageProps.title}
                  selectedProperties = { this.props.advPageProps.selectedProperties}
                  themeVariant = { this.props.advPageProps.themeVariant}
          >
          </AdvancedPageProperties>
        </div>
      </section>
    );
  }
}

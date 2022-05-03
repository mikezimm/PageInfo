import * as React from 'react';
import styles from './FpsPageInfo.module.scss';
import { IFpsPageInfoProps } from './IFpsPageInfoProps';
import { escape } from '@microsoft/sp-lodash-subset';
import PageNavigator from './PageNavigator/PageNavigator';

import ReactJson from "react-json-view";

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
          <h3>This is FPS Page Info web part</h3>
          <PageNavigator 
            description={ this.props.pageNavigator.description }
            anchorLinks={ this.props.pageNavigator.anchorLinks }
          >
          </PageNavigator>
        </div>
      </section>
    );
  }
}

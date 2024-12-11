import * as React from 'react';
import styles from './CustomBanner.module.scss';
import type { ICustomBannerProps } from './ICustomBannerProps';
//import { escape } from '@microsoft/sp-lodash-subset';

export default class CustomBanner extends React.Component<ICustomBannerProps> {
  public render(): React.ReactElement<ICustomBannerProps> {
    const {
    
      hasTeamsContext,
    
    } = this.props;

    return (
      <section className={`${styles.customBanner} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
        <img alt="Custom List Image" src={this.props.imageurl} className={styles.customImage} />

        </div>
      </section>
    );
  }
}

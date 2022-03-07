import * as React from 'react';
import styles from './SpoInterop.module.scss';
import { ISpoInteropProps } from './ISpoInteropProps';
import { escape } from '@microsoft/sp-lodash-subset';
import AppInit from './AppInit';

export default class SpoInterop extends React.Component<ISpoInteropProps, {}> {
  public render(): React.ReactElement<ISpoInteropProps> {
    const {
      description,
      isDarkTheme,
      environmentMessage,
      hasTeamsContext,
      userDisplayName,
      userLoginName,
      userEmail,
      spoContext
    } = this.props;

    return (
      <section className={`${styles.spoInterop} ${hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.welcome}>
          <img alt="" src={isDarkTheme ? require('../assets/welcome-dark.png') : require('../assets/welcome-light.png')} className={styles.welcomeImage} />
          <h2>Welcome, {escape(userDisplayName)}!</h2>
        </div>
        <div>
          <AppInit loginName={userLoginName} spoContext={spoContext} />
        </div>
      </section>
    );
  }
}

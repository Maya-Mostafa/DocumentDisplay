import * as React from 'react';
import styles from './DocumentDisplay.module.scss';
import { IDocumentDisplayProps } from './IDocumentDisplayProps';
import { escape } from '@microsoft/sp-lodash-subset';

export default function DocumentDisplay(props: IDocumentDisplayProps) {
    return (
      <section className={`${styles.documentDisplay} ${props.hasTeamsContext ? styles.teams : ''}`}>
        <div>Web part property value: <strong>{escape(props.description)}</strong></div>
        <div>
          Test
        </div>
      </section>
    );
}

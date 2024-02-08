import * as React from 'react';
import styles from './DocumentDisplay.module.scss';
import './DocumentDisplay.scss';
import { IDocumentDisplayProps } from './IDocumentDisplayProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import {getFileTypeUrl, getGraphMemberOf, isFromTargetAudience} from '../services/DataRequests';
import { Icon } from 'office-ui-fabric-react';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import { faArrowRightLong, faDownload } from '@fortawesome/free-solid-svg-icons';

export default function DocumentDisplay(props: IDocumentDisplayProps) {

  const [showBasedOnTargetAudience, setShowBasedOnTargetAudience] = React.useState(false);

  let downloadLink: string;
  if (props.filePickerResult){
    // console.log("props.filePickerResult", props.filePickerResult);
    downloadLink = props.filePickerResult.fileAbsoluteUrl;
    if (downloadLink.indexOf('?') !== -1)
      downloadLink = downloadLink.substring(0, downloadLink.indexOf('?'));
  }

  React.useEffect(()=>{
    if (props.targetAudience && props.targetAudience.length > 0){
      getGraphMemberOf(props.context).then((res: any) => {
        console.log("getGraphMemberOf res", res);
        console.log("props.targetAudience", props.targetAudience)
        setShowBasedOnTargetAudience(isFromTargetAudience(props.context, res.value, props.targetAudience));
      });
    }else{
      setShowBasedOnTargetAudience(true);
    }
  }, []);

  return (
    <>
    {showBasedOnTargetAudience &&
      <section className={`${styles.documentDisplay} ${props.hasTeamsContext ? styles.teams : ''}`}>
        <div className={styles.documentBox}>
          <div className={styles.docHeader}>
            {props.thumbnail === 'icon' &&
              <Icon iconName={props.iconPicker}/>
            }
            {props.thumbnail === 'customImg' &&
              <img height={45} src={props.customImgPicker.fileAbsoluteUrl} />
            }
            {props.thumbnail === 'fileIcon' &&
              <img height={45} src={getFileTypeUrl(props.filePickerResult && props.filePickerResult.fileName || props.documentLink)} />
            }
            <h4>
              <a rel="noreferrer" 
                target={props.openInNewTab ? "_blank" : "_self"} 
                data-interception="off" 
                href={`${props.filePickerResult && props.filePickerResult.fileAbsoluteUrl || props.documentLink}?web=1`}>
                {props.documentTitle ? 
                  <span>{props.documentTitle}</span>
                  :
                  <span>{props.filePickerResult && decodeURIComponent(props.filePickerResult.fileNameWithoutExtension)}</span>
                }
              </a>
            </h4>
          </div>
          {props.showFooter &&
            <div className={styles.docFooter}>
              {props.showMore &&
                <a rel="noreferrer" 
                  target={props.openInNewTab ? "_blank" : "_self"} 
                  data-interception="off" 
                  className={styles.moreLink} 
                  href={props.moreLink}>{props.moreTitle}
                  <FontAwesomeIcon icon={faArrowRightLong} />
                </a>
              }
              {props.showDownload && 
                // <a href={`${props.filePickerResult && props.filePickerResult.fileAbsoluteUrl.substring(0, props.filePickerResult.fileAbsoluteUrl.indexOf('?'))}?download=1`} 
                <a href={`${downloadLink}?download=1`} 
                  title='Download' 
                  download
                  className={styles.downloadLink}>
                  <FontAwesomeIcon icon={faDownload} />
                </a>
              }
            </div>
          }

        </div>
      </section>
    }
    </>
  );
}

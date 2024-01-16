import * as React from 'react';
import styles from './DocumentDisplay.module.scss';
import { IDocumentDisplayProps } from './IDocumentDisplayProps';
// import { escape } from '@microsoft/sp-lodash-subset';
import { FilePicker, IFilePickerResult } from '@pnp/spfx-controls-react/lib/FilePicker';

export default function DocumentDisplay(props: IDocumentDisplayProps) {

  //const [filePickerResult, setFilePickerResult] = React.useState(null);

  const onFilePickerSaveHandler = async (filePickerResult: IFilePickerResult[]) => {
    //setFilePickerResult(filePickerResult);
    console.log("filePickerResult", filePickerResult);
    if (filePickerResult && filePickerResult.length > 0) {
      for (let i = 0; i < filePickerResult.length; i++) {
        const item = filePickerResult[i];
        const fileResultContent = await item.downloadFileContent();
        console.log("fileResultContent", fileResultContent);
      }
    }
  }


    return (
      <section className={`${styles.documentDisplay} ${props.hasTeamsContext ? styles.teams : ''}`}>
        {/* <div>Web part property value: <strong>{escape(props.description)}</strong></div> */}
        {/* {filePickerResult} */}
        <div>
        <FilePicker
          bingAPIKey="<BING API KEY>"
          accepts= {[".gif", ".jpg", ".jpeg", ".bmp", ".dib", ".tif", ".tiff", ".ico", ".png", ".jxr", ".svg"]}
          buttonIcon="FileImage"
          onSave={(filePickerResult: IFilePickerResult[]) => onFilePickerSaveHandler(filePickerResult)}
          onChange={(filePickerResult: IFilePickerResult[]) => onFilePickerSaveHandler(filePickerResult)}
          context={props.context as any}
          buttonLabel='Upload file'
        />
        </div>
      </section>
    );
}

import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IDocumentDisplayProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;

  context: WebPartContext;
  
  documentTitle: string;
  documentLink: string;
  openInNewTab: boolean;
  browseCustomLink: boolean;

  showFooter: boolean;
  moreTitle: string;
  moreLink: string;
  showDownload: boolean;
  showMore: boolean;

  filePickerResult: any;
  iconPicker: any;
  thumbnail: any;
  customImgPicker: any;

  targetAudience: any;
}

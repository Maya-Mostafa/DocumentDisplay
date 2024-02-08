import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneChoiceGroup,
  PropertyPaneCheckbox,
  PropertyPaneToggle,
  PropertyPaneLabel,
  PropertyPaneButton,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, WebPartContext } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { PropertyFieldFilePicker, IFilePickerResult } from "@pnp/spfx-property-controls/lib/PropertyFieldFilePicker";
import { PropertyFieldIconPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldIconPicker';
import * as strings from 'DocumentDisplayWebPartStrings';
import DocumentDisplay from './components/DocumentDisplay';
import { IDocumentDisplayProps } from './components/IDocumentDisplayProps';
import { PropertyFieldPeoplePicker, IPropertyFieldGroupOrPerson, PrincipalType } from '@pnp/spfx-property-controls/lib/PropertyFieldPeoplePicker';

export interface IPropertyControlsTestWebPartProps {
  people: IPropertyFieldGroupOrPerson[];
}


export interface IDocumentDisplayWebPartProps {
  context: WebPartContext;
  description: string;
  documentTitle: string;
  showFooter: boolean;
  moreTitle: string;
  moreLink: string;
  showDownload: boolean;
  showMore: boolean;
  filePickerResult: IFilePickerResult;
  iconPicker: any;
  thumbnail: any;
  customImgPicker: any;
  documentLink: string;
  openInNewTab: boolean;
  browseCustomLink: boolean;
  targetAudience: any;
}

export default class DocumentDisplayWebPart extends BaseClientSideWebPart<IDocumentDisplayWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {
    const element: React.ReactElement<IDocumentDisplayProps> = React.createElement(
      DocumentDisplay,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userDisplayName: this.context.pageContext.user.displayName,
        
        context: this.context,

        documentTitle: this.properties.documentTitle,
        documentLink: this.properties.documentLink,

        browseCustomLink: this.properties.browseCustomLink,
        showFooter: this.properties.showFooter,
        showMore: this.properties.showMore,
        moreTitle: this.properties.moreTitle,
        moreLink: this.properties.moreLink,
        showDownload: this.properties.showDownload,

        openInNewTab : this.properties.openInNewTab,

        filePickerResult: this.properties.filePickerResult,
        iconPicker: this.properties.iconPicker,
        thumbnail: this.properties.thumbnail,
        customImgPicker: this.properties.customImgPicker,

        targetAudience: this.properties.targetAudience
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected onInit(): Promise<void> {
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              throw new Error('Unknown host');
          }

          return environmentMessage;
        });
    }

    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: any, newValue: any): void {
    console.log("onPropertyPaneFieldChanged");
    // console.log("filePickerResult", this.properties.filePickerResult);
    // console.log("thumbnail", this.properties.thumbnail);
  }

  protected onGotoSiteAssetsClick(){
    if (this.context){
      const siteAssetsUrl = `${this.context.pageContext.web.absoluteUrl}/SiteAssets` ;
      window.open(siteAssetsUrl, '_blank');
    }
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          groups: [
            {
              groupName: 'Link/Document Information',
              groupFields: [
                PropertyPaneTextField('documentTitle', {
                  label: 'Title',
                  value: this.properties.documentTitle
                }),
                PropertyPaneCheckbox('openInNewTab', {
                  checked: this.properties.openInNewTab,
                  text: 'Open link in a new tab'
                }),
                PropertyPaneToggle('browseCustomLink', {
                  checked: this.properties.browseCustomLink,
                  onText: 'Browse link from this site or your OneDrive',
                  offText: 'Type a custom link'
                }),
                PropertyFieldFilePicker('filePicker', {
                  context: this.context as any,
                  filePickerResult: this.properties.filePickerResult,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (e: IFilePickerResult) => { console.log(e); this.properties.filePickerResult = e;  },
                  onChanged: (e: IFilePickerResult) => { console.log(e); this.properties.filePickerResult = e; },
                  key: "filePickerId",
                  buttonLabel: "Browse Link",
                  hideLinkUploadTab: true,
                  hideLocalUploadTab : true,
                  hideStockImages: true,
                  hideWebSearchTab: true,
                  allowExternalLinks: false,
                  disabled: !this.properties.browseCustomLink,
                }),
                PropertyPaneTextField('documentLink', {
                  label: 'Custom Link',
                  value: this.properties.documentLink,
                  disabled: this.properties.browseCustomLink
                }),
              ]
            },
            {
              groupName: 'Footer Option',
              groupFields: [
                PropertyPaneCheckbox('showFooter', {
                  checked: this.properties.showFooter,
                  text: 'Show footer'
                }),
                PropertyPaneCheckbox('showMore', {
                  checked: this.properties.showMore,
                  text: 'Show the "more" option'
                }),
                PropertyPaneTextField('moreTitle', {
                  label: 'More Text',
                  value: this.properties.moreTitle
                }),
                PropertyPaneTextField('moreLink', {
                  label: 'More Link',
                  value: this.properties.moreLink
                }),
                PropertyPaneCheckbox('showDownload', {
                  checked: this.properties.showDownload,
                  text: 'Show the "download" option'
                }),
              ]
            },
            {
              groupName: 'Document Icon',
              groupFields: [
                PropertyPaneChoiceGroup('thumbnail', {
                  label: 'Thumbnail',
                  options : [
                    {key: 'icon', text:'Icon'}, 
                    {key: 'customImg', text:'Custom Image'}, 
                    {key: 'fileIcon', text:'File Icon'}, 
                    {key: 'none', text:'None', checked: true}
                  ],
                }),
                PropertyFieldIconPicker('iconPicker', {
                  // currentIcon: this.properties.iconPicker,
                  key: "iconPickerId",
                  onSave: (icon: string) => { console.log(icon); this.properties.iconPicker = icon; },
                  onChanged:(icon: string) => { console.log(icon);  },
                  buttonLabel: "Change",
                  renderOption: "panel",
                  properties: this.properties,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  buttonClassName:  'iconPicker-' + this.properties.thumbnail ,     
                }),
                PropertyFieldFilePicker('customImgPicker', {
                  buttonIcon:'',
                  context: this.context as any,
                  filePickerResult: this.properties.customImgPicker,
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  properties: this.properties,
                  onSave: (e: IFilePickerResult) => { console.log(e); this.properties.customImgPicker = e;  },
                  onChanged: (e: IFilePickerResult) => { console.log(e); this.properties.customImgPicker = e; },
                  key: "customImgPickerId",
                  buttonLabel: "Change",
                  buttonClassName: 'customImgPicker-btn filePicker-' + this.properties.thumbnail,
                  hideLocalUploadTab : true,
                  hideLinkUploadTab: true,
                  allowExternalLinks: true,
                  accepts: [".gif", ".jpg", ".jpeg", ".bmp", ".dib", ".tif", ".tiff", ".ico", ".png", ".jxr", ".svg"]
                }),
                PropertyPaneLabel('customImgNote', {
                  text: 'To upload a custom image to site assests please use the button below'
                }),
                PropertyPaneButton('goToSiteAssetsBtn', {
                  text: 'Go to Site Assets',
                  onClick: this.onGotoSiteAssetsClick.bind(this)
                })
              ]
            },
            {
              groupName: 'Target Audience',
              groupFields: [
                PropertyFieldPeoplePicker('targetAudience', {
                  label: 'Target Audience e.g. User(s), Group(s)',
                  initialData: this.properties.targetAudience,
                  allowDuplicate: false,
                  principalType: [PrincipalType.Users, PrincipalType.SharePoint, PrincipalType.Security],
                  onPropertyChange: this.onPropertyPaneFieldChanged.bind(this),
                  context: this.context as any,
                  properties: this.properties,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  key: 'peopleFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}

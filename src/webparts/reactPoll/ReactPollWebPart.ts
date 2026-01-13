import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  IPropertyPaneGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'ReactPollWebPartStrings';
import ReactPoll from './components/ReactPoll';
import { IReactPollProps } from './components/IReactPollProps';
import { Language, IPollTextConfiguration, defaultTextConfiguration } from '../../models/models';

/**
 * Web Part Properties Interface
 * Includes language selection and bilingual text configuration
 */
export interface IReactPollWebPartProps {
  description: string;
  // Language Selection - Controls RTL/LTR and text display
  language: Language;
  // English Text Properties
  webPartTitleEn: string;
  submitButtonTextEn: string;
  loadingMessageEn: string;
  noQuestionsMessageEn: string;
  alreadyVotedMessageEn: string;
  selectOptionMessageEn: string;
  thankYouMessageEn: string;
  resultsTitleEn: string;
  // Arabic Text Properties
  webPartTitleAr: string;
  submitButtonTextAr: string;
  loadingMessageAr: string;
  noQuestionsMessageAr: string;
  alreadyVotedMessageAr: string;
  selectOptionMessageAr: string;
  thankYouMessageAr: string;
  resultsTitleAr: string;
}

export default class ReactPollWebPart extends BaseClientSideWebPart<IReactPollWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  /**
   * Initialize default property values for backward compatibility
   * English is the default language
   */
  protected onInit(): Promise<void> {
    // Set default values if not already set (backward compatibility)
    if (!this.properties.language) {
      this.properties.language = Language.English;
    }
    // Initialize English text defaults
    if (!this.properties.webPartTitleEn) this.properties.webPartTitleEn = defaultTextConfiguration.webPartTitle.en;
    if (!this.properties.submitButtonTextEn) this.properties.submitButtonTextEn = defaultTextConfiguration.submitButtonText.en;
    if (!this.properties.loadingMessageEn) this.properties.loadingMessageEn = defaultTextConfiguration.loadingMessage.en;
    if (!this.properties.noQuestionsMessageEn) this.properties.noQuestionsMessageEn = defaultTextConfiguration.noQuestionsMessage.en;
    if (!this.properties.alreadyVotedMessageEn) this.properties.alreadyVotedMessageEn = defaultTextConfiguration.alreadyVotedMessage.en;
    if (!this.properties.selectOptionMessageEn) this.properties.selectOptionMessageEn = defaultTextConfiguration.selectOptionMessage.en;
    if (!this.properties.thankYouMessageEn) this.properties.thankYouMessageEn = defaultTextConfiguration.thankYouMessage.en;
    if (!this.properties.resultsTitleEn) this.properties.resultsTitleEn = defaultTextConfiguration.resultsTitle.en;
    // Initialize Arabic text defaults
    if (!this.properties.webPartTitleAr) this.properties.webPartTitleAr = defaultTextConfiguration.webPartTitle.ar;
    if (!this.properties.submitButtonTextAr) this.properties.submitButtonTextAr = defaultTextConfiguration.submitButtonText.ar;
    if (!this.properties.loadingMessageAr) this.properties.loadingMessageAr = defaultTextConfiguration.loadingMessage.ar;
    if (!this.properties.noQuestionsMessageAr) this.properties.noQuestionsMessageAr = defaultTextConfiguration.noQuestionsMessage.ar;
    if (!this.properties.alreadyVotedMessageAr) this.properties.alreadyVotedMessageAr = defaultTextConfiguration.alreadyVotedMessage.ar;
    if (!this.properties.selectOptionMessageAr) this.properties.selectOptionMessageAr = defaultTextConfiguration.selectOptionMessage.ar;
    if (!this.properties.thankYouMessageAr) this.properties.thankYouMessageAr = defaultTextConfiguration.thankYouMessage.ar;
    if (!this.properties.resultsTitleAr) this.properties.resultsTitleAr = defaultTextConfiguration.resultsTitle.ar;

    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
    });
  }

  /**
   * Build the text configuration object based on current property values
   */
  private _getTextConfiguration(): IPollTextConfiguration {
    return {
      webPartTitle: { en: this.properties.webPartTitleEn, ar: this.properties.webPartTitleAr },
      submitButtonText: { en: this.properties.submitButtonTextEn, ar: this.properties.submitButtonTextAr },
      loadingMessage: { en: this.properties.loadingMessageEn, ar: this.properties.loadingMessageAr },
      noQuestionsMessage: { en: this.properties.noQuestionsMessageEn, ar: this.properties.noQuestionsMessageAr },
      alreadyVotedMessage: { en: this.properties.alreadyVotedMessageEn, ar: this.properties.alreadyVotedMessageAr },
      selectOptionMessage: { en: this.properties.selectOptionMessageEn, ar: this.properties.selectOptionMessageAr },
      thankYouMessage: { en: this.properties.thankYouMessageEn, ar: this.properties.thankYouMessageAr },
      resultsTitle: { en: this.properties.resultsTitleEn, ar: this.properties.resultsTitleAr }
    };
  }

  public render(): void {
    const element: React.ReactElement<IReactPollProps> = React.createElement(
      ReactPoll,
      {
        description: this.properties.description,
        isDarkTheme: this._isDarkTheme,
        environmentMessage: this._environmentMessage,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        userEmail: this.context.pageContext.user.email,
        context: this.context,
        webServerRelativeUrl: this.context.pageContext.web.serverRelativeUrl,
        // Bilingual RTL/LTR support properties
        language: this.properties.language || Language.English,
        textConfiguration: this._getTextConfiguration()
      }
    );

    ReactDom.render(element, this.domElement);
  }



  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
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

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    // Determine which text fields to show based on selected language
    const isArabic = this.properties.language === Language.Arabic;
    const langSuffix = isArabic ? 'Ar' : 'En';

    // Build text configuration group - dynamically shows English or Arabic fields
    const textConfigGroup: IPropertyPaneGroup = {
      groupName: strings.TextConfigGroupName,
      groupFields: [
        PropertyPaneTextField(`webPartTitle${langSuffix}`, {
          label: strings.WebPartTitleLabel
        }),
        PropertyPaneTextField(`submitButtonText${langSuffix}`, {
          label: strings.SubmitButtonLabel
        }),
        PropertyPaneTextField(`loadingMessage${langSuffix}`, {
          label: strings.LoadingMessageLabel
        }),
        PropertyPaneTextField(`noQuestionsMessage${langSuffix}`, {
          label: strings.NoQuestionsMessageLabel
        }),
        PropertyPaneTextField(`selectOptionMessage${langSuffix}`, {
          label: strings.SelectOptionMessageLabel
        }),
        PropertyPaneTextField(`thankYouMessage${langSuffix}`, {
          label: strings.ThankYouMessageLabel
        }),
        PropertyPaneTextField(`resultsTitle${langSuffix}`, {
          label: strings.ResultsTitleLabel
        })
      ]
    };

    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.LanguageGroupName,
              groupFields: [
                PropertyPaneDropdown('language', {
                  label: strings.LanguageFieldLabel,
                  options: [
                    { key: Language.English, text: 'English' },
                    { key: Language.Arabic, text: 'العربية (Arabic)' }
                  ],
                  selectedKey: this.properties.language || Language.English
                })
              ]
            }
          ]
        },
        {
          header: {
            description: isArabic ? strings.ArabicTextDescription : strings.EnglishTextDescription
          },
          groups: [textConfigGroup]
        }
      ]
    };
  }
}

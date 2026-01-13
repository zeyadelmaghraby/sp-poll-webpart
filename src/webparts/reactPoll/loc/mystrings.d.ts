declare interface IReactPollWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  // Language Settings
  LanguageGroupName: string;
  LanguageFieldLabel: string;
  // English Text Group
  EnglishTextGroupName: string;
  EnglishTextDescription: string;
  // Arabic Text Group
  ArabicTextGroupName: string;
  ArabicTextDescription: string;
  // Generic Text Config Group
  TextConfigGroupName: string;
  // Text Field Labels
  WebPartTitleLabel: string;
  SubmitButtonLabel: string;
  LoadingMessageLabel: string;
  NoQuestionsMessageLabel: string;
  SelectOptionMessageLabel: string;
  ThankYouMessageLabel: string;
  ResultsTitleLabel: string;
  // Environment Messages
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
}

declare module 'ReactPollWebPartStrings' {
  const strings: IReactPollWebPartStrings;
  export = strings;
}

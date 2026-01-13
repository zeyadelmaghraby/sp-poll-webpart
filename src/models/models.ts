/**
 * Supported languages for the Poll Web Part
 * Used for bilingual (English/Arabic) support with RTL/LTR layout switching
 */
export enum Language {
  English = 'en',
  Arabic = 'ar'
}

/**
 * Bilingual text configuration interface
 * Stores separate text values for English and Arabic
 */
export interface IBilingualText {
  en: string;
  ar: string;
}

/**
 * Configurable UI text properties for the Poll Web Part
 * All user-facing text is configurable from the Property Pane
 */
export interface IPollTextConfiguration {
  // Web Part Title
  webPartTitle: IBilingualText;
  // Submit Button
  submitButtonText: IBilingualText;
  // Loading Message
  loadingMessage: IBilingualText;
  // No Questions Message
  noQuestionsMessage: IBilingualText;
  // Already Voted Message
  alreadyVotedMessage: IBilingualText;
  // Select Option Validation Message
  selectOptionMessage: IBilingualText;
  // Thank You Message
  thankYouMessage: IBilingualText;
  // Results Title
  resultsTitle: IBilingualText;
}

/**
 * Default text configuration with English and Arabic translations
 */
export const defaultTextConfiguration: IPollTextConfiguration = {
  webPartTitle: {
    en: 'Poll',
    ar: 'استطلاع'
  },
  submitButtonText: {
    en: 'Submit',
    ar: 'إرسال'
  },
  loadingMessage: {
    en: 'Loading poll...',
    ar: 'جاري تحميل الاستطلاع...'
  },
  noQuestionsMessage: {
    en: 'No active poll questions available.',
    ar: 'لا توجد أسئلة استطلاع نشطة متاحة.'
  },
  alreadyVotedMessage: {
    en: 'You have already voted on this poll.',
    ar: 'لقد قمت بالتصويت بالفعل في هذا الاستطلاع.'
  },
  selectOptionMessage: {
    en: 'Please select an option to vote.',
    ar: 'يرجى اختيار خيار للتصويت.'
  },
  thankYouMessage: {
    en: 'Thank you for your vote!',
    ar: 'شكراً لتصويتك!'
  },
  resultsTitle: {
    en: 'Poll Results',
    ar: 'نتائج الاستطلاع'
  }
};

export interface IQuestion {
  id: number;
  question: string;
  options:IOption[];
  answer?:IAnswer;
  selectedOption?:string;
  
}

export interface IOption {
  key: string;
  text: string;
}

export interface IAnswer{
  allAnswers:number[];
  isCurrentUserAnswered:boolean;
}

export interface ISelectedOption{
  selectedQuestionId:number;
  selectedOption:string;
}

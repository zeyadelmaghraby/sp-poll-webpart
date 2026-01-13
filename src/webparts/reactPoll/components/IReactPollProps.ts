import { WebPartContext } from '@microsoft/sp-webpart-base';
import { Language, IPollTextConfiguration } from '../../../models/models';

/**
 * Props interface for the ReactPoll component
 * Includes bilingual (English/Arabic) support with RTL/LTR layout switching
 */
export interface IReactPollProps {
  description: string;
  isDarkTheme: boolean;
  environmentMessage: string;
  hasTeamsContext: boolean;
  userEmail: string;
  context: WebPartContext;
  webServerRelativeUrl: string;
  /**
   * Selected language for the Web Part
   * Controls text display and RTL/LTR direction
   */
  language: Language;
  /**
   * Configurable text values for both English and Arabic
   * All UI text is stored and configured separately for each language
   */
  textConfiguration: IPollTextConfiguration;
}

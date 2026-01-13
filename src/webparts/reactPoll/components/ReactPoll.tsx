import * as React from "react";
import styles from "./ReactPoll.module.scss";
import { IReactPollProps } from "./IReactPollProps";
import { useGetData } from "../../../apiHooks/useGetData";
import { useEffect, useMemo } from "react";
import {
  ChoiceGroup,
  IChoiceGroupOption,
} from "@fluentui/react/lib/ChoiceGroup";
import { PrimaryButton } from "@fluentui/react/lib/Button";
import { IQuestion, Language, IBilingualText } from "../../../models/models";
import {
  ChartControl,
  ChartType,
} from "@pnp/spfx-controls-react/lib/ChartControl";
import { ShimmerLoadder } from "./Loader";

/**
 * ReactPoll Component
 * A bilingual (English/Arabic) poll web part with RTL/LTR support
 * 
 * Features:
 * - Dynamic language switching between English and Arabic
 * - Automatic RTL/LTR layout based on selected language
 * - Configurable UI text from Property Pane
 * - Responsive design for all device sizes
 */
const ReactPoll: React.FunctionComponent<IReactPollProps> = (props) => {
  const { language, textConfiguration } = props;
  
  const {
    isLoading,
    isSubmitting,
    questions,
    getQuestion,
    submitAnswer,
    setQuestions,
  } = useGetData(props.context, props.userEmail, props.webServerRelativeUrl);

  /**
   * Helper function to get text based on current language
   * Returns English text by default for backward compatibility
   */
  const getText = useMemo(() => {
    return (bilingualText: IBilingualText): string => {
      return language === Language.Arabic ? bilingualText.ar : bilingualText.en;
    };
  }, [language]);

  /**
   * Determine if the current language is RTL (Arabic)
   * Used for applying RTL/LTR CSS classes and direction attribute
   */
  const isRTL = useMemo(() => language === Language.Arabic, [language]);

  /**
   * Get the appropriate direction class based on language
   * 'rtl' for Arabic, 'ltr' for English (default)
   */
  const directionClass = useMemo(() => isRTL ? styles.rtl : styles.ltr, [isRTL]);

  /**
   * Get the direction attribute value for HTML elements
   * 'rtl' for Arabic, 'ltr' for English
   */
  const directionAttr = useMemo(() => isRTL ? 'rtl' : 'ltr', [isRTL]);

  useEffect(() => {
    // Get Question on component mount
    getQuestion()
      .then(() => {
        return;
      })
      .catch((e) => {
        console.log(e);
      });
  }, []);

  /**
   * Handle submit button click
   * Submits the user's answer for the selected option
   */
  const handleSubmitClick = (question: IQuestion): void => {
    submitAnswer(question.id, question.selectedOption, props.userEmail)
      .then(() => {
        return;
      })
      .catch((e) => {
        console.log(e);
      });
  };

  /**
   * Handle option selection change
   * Updates the selected option for the question
   */
  const changeSelectedOption = (
    ev: React.FormEvent<HTMLElement | HTMLInputElement>,
    option: IChoiceGroupOption,
    questionId: number
  ): void => {
    const updatedQuestions = questions.map((question: IQuestion) => {
      if (question.id === questionId) {
        return { ...question, selectedOption: option.key };
      }
      return question;
    });
    setQuestions(updatedQuestions);
  };

  return (
    // Main container with RTL/LTR direction class applied
    <div className={`${styles.reactPoll} ${directionClass}`} dir={directionAttr}>
      {/* Web Part Title - Configurable from Property Pane */}
      <div className={styles.pollTitle}>
        {getText(textConfiguration.webPartTitle)}
      </div>

      {/* Show Loader while Loading Question */}
      {isLoading && (
        <div className={styles.loadingMessage}>
          <ShimmerLoadder />
          <p>{getText(textConfiguration.loadingMessage)}</p>
        </div>
      )}

      {/* Show "No Questions" message when no active polls */}
      {!isLoading && (!questions || questions.length === 0) && (
        <div className={styles.noQuestionsMessage}>
          {getText(textConfiguration.noQuestionsMessage)}
        </div>
      )}

      {/* Render Poll Questions */}
      {questions &&
        questions.length > 0 &&
        questions.map((question) => (
          <div key={question.id}>
            {/* Show Chart if Current User has already voted */}
            {question.answer.isCurrentUserAnswered && (
              <React.Fragment key={question.id}>
                {/* Results container with dynamic direction */}
                <div className={styles["ms-Grid"]} dir={directionAttr}>
                  {/* Results Title */}
                  <div className={styles.resultsTitle}>
                    {getText(textConfiguration.resultsTitle)}
                  </div>
                  
                  <div className={styles["ms-Grid-row"]}>
                    <div
                      className={`${styles["ms-Grid-col"]} ${styles["ms-sm12"]} ${styles["ms-font-m-plus"]} ${styles["ms-fontWeight-bold"]} ${styles.questionLabel}`}
                    >
                      {question.question}
                    </div>
                  </div>
                  <div className={`${styles["ms-Grid-row"]} ${styles.chartRow}`}>
                    <div
                      className={`${styles["ms-Grid-col"]} ${styles["ms-sm12"]} ${styles["ms-md12"]} ${styles["ms-lg6"]} ${styles.chartContainer}`}
                    >
                      <ChartControl
                        type={ChartType.Pie}
                        data={{
                          labels: question.options
                            .filter((i) => i !== null)
                            .map((i) => i?.text),
                          datasets: [
                            {
                              label: question.question,
                              data: question.answer.allAnswers,
                            },
                          ],
                        }}
                        options={{
                          Animation: false,
                          legend: {
                            display: true,
                            // Position legend based on language direction
                            position: isRTL ? "left" : "right",
                            rtl: isRTL, // Enable RTL for chart legend
                          },
                          title: {
                            display: false,
                          },
                        }}
                      />
                    </div>
                  </div>
                </div>
              </React.Fragment>
            )}

            {/* Show Question with Options if Current User hasn't voted yet */}
            {!question.answer.isCurrentUserAnswered && (
              <React.Fragment key={question.id}>
                <div className={styles["ms-Grid"]} dir={directionAttr}>
                  <div className={styles["ms-Grid-row"]}>
                    <div className={`${styles["ms-Grid-col"]} ${styles["ms-sm12"]}`}>
                      <ChoiceGroup
                        options={question.options.filter((i) => i !== null)}
                        label={question.question}
                        onChange={(e, selectedOption) =>
                          changeSelectedOption(e, selectedOption, question.id)
                        }
                        // Apply RTL styling to ChoiceGroup
                        styles={{
                          label: {
                            textAlign: isRTL ? 'right' : 'left',
                            fontFamily: isRTL 
                              ? "'Segoe UI', 'Arabic Typesetting', 'Traditional Arabic', Tahoma, Arial, sans-serif"
                              : undefined,
                          },
                          flexContainer: {
                            direction: isRTL ? 'rtl' : 'ltr',
                          }
                        }}
                      />
                    </div>
                  </div>
                  <div className={`${styles["ms-Grid-row"]} ${styles.submitButtonRow}`}>
                    <div className={`${styles["ms-Grid-col"]} ${styles["ms-sm12"]}`}>
                      <PrimaryButton
                        text={getText(textConfiguration.submitButtonText)}
                        disabled={
                          question.selectedOption === undefined || isSubmitting
                        }
                        onClick={() => handleSubmitClick(question)}
                        // Apply RTL styling to button if needed
                        styles={{
                          root: {
                            minWidth: '120px',
                          },
                          label: {
                            fontFamily: isRTL 
                              ? "'Segoe UI', 'Arabic Typesetting', 'Traditional Arabic', Tahoma, Arial, sans-serif"
                              : undefined,
                          }
                        }}
                      />
                    </div>
                  </div>
                </div>
              </React.Fragment>
            )}
          </div>
        ))}
    </div>
  );
};

export default ReactPoll;

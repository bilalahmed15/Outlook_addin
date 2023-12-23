/* eslint-disable no-undef */
import * as React from "react";
import { DefaultButton } from "@fluentui/react";
import Progress from "./Progress";
import { ChatCompletionRequestMessage, Configuration, OpenAIApi } from "openai";

/* global require */

export interface AppProps {
  title: string;
  isOfficeInitialized: boolean;
}

export interface AppState {
  generatedText: string;
  startText: string;
  finalMailText: string;
  isLoading: boolean;
  summary: string;
  rawReply: string; // Add this new state variable
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props) {
    super(props);
  
    this.state = {
      generatedText: "",
      startText: "",
      finalMailText: "",
      isLoading: false,
      summary: "",
      rawReply: "", // Initialize the new state variable
    };
  }

  componentDidMount() {
    if (this.props.isOfficeInitialized) {
      this.onSummarize();
    }
  }

  generateText = async () => {
    var current = this;
    const configuration = new Configuration({
      apiKey: "",
    });
    const openai = new OpenAIApi(configuration);
    current.setState({ isLoading: true });
  
    // The original email content and the user's raw reply
    const originalEmailContent = this.state.startText;
  
    const userRawReply = this.state.rawReply;
    const response = await openai.createChatCompletion({
      model: "gpt-3.5-turbo",
      messages: [
        {
          role: "system",
          content: "You are a helpful assistant. Refine the user's raw reply into a professional response suitable as a reply to the original email.",
        },
        { 
          role: "user", 
          content: `Original email: ${originalEmailContent} User's raw reply: ${userRawReply}. Please refine this into a professional response.` 
        },
      ],
    });
    current.setState({ isLoading: false });
    current.setState({ generatedText: response.data.choices[0].message.content });
  };
  

  insertIntoMail = () => {
    const finalText = this.state.finalMailText.length === 0 ? this.state.generatedText : this.state.finalMailText;
    Office.context.mailbox.item.body.setSelectedDataAsync(finalText, {
      coercionType: Office.CoercionType.Text,
    });
  };

  onSummarize = async () => {
    try {
      this.setState({ isLoading: true });
      const summary = await this.summarizeMail();
      this.setState({ summary: summary, isLoading: false });
    } catch (error) {
      this.setState({ summary: error, isLoading: false });
    }
  };

  summarizeMail(): Promise<any> {
    return new Office.Promise(function (resolve, reject) {
      try {
        Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async function (asyncResult) {
          const configuration = new Configuration({
            apiKey: "",
          });
          const openai = new OpenAIApi(configuration);

          const mailText = asyncResult.value.split(" ").slice(0, 800).join(" ");

          const messages: ChatCompletionRequestMessage[] = [
            {
              role: "system",
              content:
                "You are a helpful assistant that can help users to better manage emails. The mail thread can be made by multiple prompts.",
            },
            {
              role: "user",
              content: "Summarize the following mail thread and summarize it with a bullet list: " + mailText,
            },
          ];

          const response = await openai.createChatCompletion({
            model: "gpt-3.5-turbo",
            messages: messages,
          });

          resolve(response.data.choices[0].message.content);
        });
      } catch (error) {
        reject(error);
      }
    });
  }

  ProgressSection = () => {
    if (this.state.isLoading) {
      return <Progress title="Loading..." message="The AI is working..." />;
    } else {
      return <> </>;
    }
  };

  BusinessMailSection = () => {
    return (
      <>
         <p>Type your raw email reply here:</p>
      <textarea
        className="ms-welcome"
        onChange={(e) => this.setState({ rawReply: e.target.value })}
        rows={5}
        cols={40}
      />
        <p>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.generateText} // Trigger reply generation
          >
            Generate Structured Reply
          </DefaultButton>
        </p>
        <this.ProgressSection />

        
        <textarea
          className="ms-welcome"
          defaultValue={this.state.generatedText}
          onChange={(e) => this.setState({ finalMailText: e.target.value })}
          rows={15}
          cols={40}
        />
        <p>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.insertIntoMail}
          >
            Insert into mail
          </DefaultButton>
        </p>
      </>
    );
  };

  render() {
    const { title, isOfficeInitialized } = this.props;

    if (!isOfficeInitialized) {
      return (
        <Progress
          title={title}
          logo={require("./../../../assets/logo-filled.png")}
          message="Please sideload your addin to see app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <main className="ms-welcome__main">
          <h2 className="ms-font-xl ms-fontWeight-semilight ms-fontColor-neutralPrimary ms-u-slideUpIn20">
            Outlook AI Assistant
          </h2>

          
          <div>
            <this.BusinessMailSection />
          </div>
          <div>
            <p>Summarized Mail:</p>
            <textarea className="ms-welcome" value={this.state.summary} readOnly rows={15} cols={40} />
          </div>
        </main>
      </div>
    );
  }
}

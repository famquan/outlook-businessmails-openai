/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
import React from "react";
import { DefaultButton } from "@fluentui/react";
import Progress from "./Progress";
import OpenAI from "openai";

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
  isGenerateBusinessMailActive: boolean;
  isSummarizeMailActive: boolean;
  summary: string;
}

export default class App extends React.Component<AppProps, AppState> {
  constructor(props: AppProps) {
    super(props);

    const url = window.location.href;
    const isGenerateBusinessMailActive = url.includes("compose");
    const isSummarizeMailActive = url.includes("summary");

    this.state = {
      generatedText: "",
      startText: "",
      finalMailText: "",
      isLoading: false,
      isGenerateBusinessMailActive,
      isSummarizeMailActive,
      summary: "",
    };
  }

  // Reusable Progress Section
  ProgressSection = () => (
    this.state.isLoading ? <Progress title="Loading..." message="The AI is working..." /> : null
  );

  // Reusable TextArea Component
  TextArea = ({ value, onChange, rows, cols }: { value: string; onChange: React.ChangeEventHandler<HTMLTextAreaElement>; rows: number; cols: number }) => (
    <textarea
      className="ms-welcome"
      value={value}
      onChange={onChange}
      rows={rows}
      cols={cols}
    />
  );

  showGenerateBusinessMail = () => {
    this.setState({ isGenerateBusinessMailActive: true, isSummarizeMailActive: false });
  };

  showSummarizeMail = () => {
    this.setState({ isGenerateBusinessMailActive: false, isSummarizeMailActive: true });
  };

  generateText = async () => {
    this.setState({ isLoading: true });
    const openai = new OpenAI({ 
      apiKey: process.env.OPENAI_API_KEY,
      dangerouslyAllowBrowser: true
    });

    try {
      const response = await openai.chat.completions.create({
        model: "gpt-4o-mini",
        messages: [
          { role: "system", content: "You are a helpful assistant that can help users create professional business content." },
          { role: "user", content: `Turn the following text into a professional business mail: ${this.state.startText}` },
        ],
        max_tokens: 150,
      });
      this.setState({ generatedText: response.choices[0]?.message?.content || "", isLoading: false });
    } catch (error) {
      console.error("OpenAI API Error:", error);
      this.setState({ isLoading: false });
    }
  };

  insertIntoMail = () => {
    const finalText = this.state.finalMailText || this.state.generatedText;
    Office.context.mailbox.item.body.setSelectedDataAsync(finalText, { coercionType: Office.CoercionType.Text });
  };

  onSummarize = async () => {
    try {
      this.setState({ isLoading: true });
      const summary = await this.summarizeMail();
      this.setState({ summary, isLoading: false });
    } catch (error) {
      console.error("Summarize Error:", error);
      this.setState({ summary: `Error: ${error.message || error}`, isLoading: false });
    }
  };

  summarizeMail = (): Promise<string> => {
    return new Promise((resolve, reject) => {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async (asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
          return reject(new Error("Failed to retrieve email body."));
        }

        try {
          const openai = new OpenAI({ 
            apiKey: process.env.OPENAI_API_KEY,
            dangerouslyAllowBrowser: true
          });
          const mailText = asyncResult.value.split(" ").slice(0, 800).join(" ");

          const response = await openai.chat.completions.create({
            model: "gpt-4o-mini",
            messages: [
              { role: "system", content: "You are a helpful assistant to help users better manage emails." },
              { role: "user", content: `Summarize the following mail thread into a bullet list: ${mailText}` },
            ],
            max_tokens: 150,
          });

          resolve(response.choices[0]?.message?.content || "No summary available.");
        } catch (error) {
          reject(new Error(`OpenAI API Error: ${error.message || error}`));
        }
      });
    });
  };

  BusinessMailSection = () => {
    if (!this.state.isGenerateBusinessMailActive) return null;

    return (
      <>
        <p>Briefly describe what you want to communicate in the mail:</p>
        <this.TextArea
          value={this.state.startText}
          onChange={(e) => this.setState({ startText: e.target.value })}
          rows={5}
          cols={40}
        />
        <DefaultButton
          className="ms-welcome__action"
          iconProps={{ iconName: "ChevronRight" }}
          onClick={this.generateText}
        >
          Generate text
        </DefaultButton>
        <this.ProgressSection />
        <this.TextArea
          value={this.state.generatedText}
          onChange={(e) => this.setState({ finalMailText: e.target.value })}
          rows={15}
          cols={40}
        />
        <DefaultButton
          className="ms-welcome__action"
          iconProps={{ iconName: "ChevronRight" }}
          onClick={this.insertIntoMail}
        >
          Insert into mail
        </DefaultButton>
      </>
    );
  };

  SummarizeMailSection = () => {
    if (!this.state.isSummarizeMailActive) return null;

    return (
      <>
        <p>Summarize mail</p>
        <DefaultButton
          className="ms-welcome__action"
          iconProps={{ iconName: "ChevronRight" }}
          onClick={this.onSummarize}
        >
          Summarize
        </DefaultButton>
        <this.ProgressSection />
        <this.TextArea
          value={this.state.summary}
          onChange={() => {}}
          rows={15}
          cols={40}
        />
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
          message="Please sideload your add-in to see the app body."
        />
      );
    }

    return (
      <div className="ms-welcome">
        <main className="ms-welcome__main">
          <h2>Outlook AI Assistant</h2>
          <p>Choose your service:</p>
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.showSummarizeMail}
          >Summarize it!</DefaultButton>
          <this.BusinessMailSection />
          <this.SummarizeMailSection />
          <DefaultButton
            className="ms-welcome__action"
            iconProps={{ iconName: "ChevronRight" }}
            onClick={this.showGenerateBusinessMail}
          >Reply it!</DefaultButton>
        </main>
      </div>
    );
  }
}

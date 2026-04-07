export interface AIClientConfig {
  apiKey: string;
  baseUrl: string;
  model: string;
}

interface ChatMessage {
  role: "system" | "user" | "assistant";
  content: string;
}

interface ChatCompletionResponse {
  choices: Array<{
    message: {
      content: string;
    };
  }>;
}

const SYSTEM_PROMPT =
  "You are a professional email assistant. You help users manage their email correspondence efficiently and professionally. " +
  "Your responses are clear, concise, and appropriate for a business context.";

export class AIClient {
  private apiKey: string;
  private baseUrl: string;
  private model: string;

  constructor(config: AIClientConfig) {
    this.apiKey = config.apiKey;
    this.baseUrl = config.baseUrl.replace(/\/$/, "");
    this.model = config.model;
  }

  private async chat(messages: ChatMessage[]): Promise<string> {
    const url = `${this.baseUrl}/chat/completions`;

    const response = await fetch(url, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${this.apiKey}`,
      },
      body: JSON.stringify({
        model: this.model,
        messages,
        temperature: 0.7,
      }),
    });

    if (!response.ok) {
      const errorText = await response.text();
      throw new Error(`API request failed (${response.status}): ${errorText}`);
    }

    const data: ChatCompletionResponse = await response.json();

    if (!data.choices || data.choices.length === 0) {
      throw new Error("No response returned from API");
    }

    return data.choices[0].message.content;
  }

  async summarize(emailText: string): Promise<string> {
    const messages: ChatMessage[] = [
      { role: "system", content: SYSTEM_PROMPT },
      {
        role: "user",
        content:
          `Please provide a concise summary of the following email thread. ` +
          `Highlight key points, action items, and any deadlines mentioned.\n\n${emailText}`,
      },
    ];
    return this.chat(messages);
  }

  async generateReply(emailText: string): Promise<string> {
    const messages: ChatMessage[] = [
      { role: "system", content: SYSTEM_PROMPT },
      {
        role: "user",
        content:
          `Please draft a professional reply to the following email. ` +
          `The reply should be courteous, clear, and address all points raised.\n\n${emailText}`,
      },
    ];
    return this.chat(messages);
  }

  async improveText(draftText: string): Promise<string> {
    const messages: ChatMessage[] = [
      { role: "system", content: SYSTEM_PROMPT },
      {
        role: "user",
        content:
          `Please rewrite the following email draft to be clearer, more professional, and more concise ` +
          `while preserving the original intent and key information.\n\n${draftText}`,
      },
    ];
    return this.chat(messages);
  }
}

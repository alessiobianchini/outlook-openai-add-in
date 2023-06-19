/* global global, Office, self, window */

import { Configuration, OpenAIApi } from "openai";

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});


function action(event: Office.AddinCommands.Event) {
  var setting = Office.context.roamingSettings.get('openApiToken');
  if (!setting) {
    Office.context.mailbox.item.setSelectedDataAsync("OpenAI token not configured \r\n", { coercionType: Office.CoercionType.Text });
    event.completed();
  }
  else {
    try {
      Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, async function (asyncResult) {
        const configuration = new Configuration({
          apiKey: Office.context.roamingSettings.get('openApiToken'),
        });
        const openai = new OpenAIApi(configuration);
        var content = "";
        var data = asyncResult.value?.data;
        const endsWithNewline = data.endsWith("\r") || data.endsWith("\n") || data.endsWith("\r\n");

        if (event.source.id == 'GenerateBusinessMail') {
          content = `Could you generate a business email based on the following text, preserving its original language? ${data}`;
        }
        else if (event.source.id == 'TranslateToEnglish') {
          content = "Can you translate to english the followng text, preserving the layout? " + data;
        }
        else if (event.source.id == 'CorrectGrammar') {
          content = "Can you correct spelling and grammar of the followng text, preserving it's original language? " + data;
        }
        const response = await openai.createChatCompletion({
          model: "gpt-3.5-turbo",
          messages: [
            {
              role: "system",
              content: "You are a helpful assistant.",
            },
            {
              role: "user",
              content: content,
            },
          ],
        });

        let res = response.data.choices[0].message.content;
        res = res?.replace(/(^"|"$)/g, '');
        const resEndsWithNewline = res.endsWith("\r");

        if (endsWithNewline && !resEndsWithNewline) {
          res += '\r';
        }

        Office.context.mailbox.item.setSelectedDataAsync(res, { coercionType: Office.CoercionType.Text });
        event.completed();
      });
    } catch (error) {
      Office.context.mailbox.item.setSelectedDataAsync(`Failed to run ${event.source.id} action \r\n`, { coercionType: Office.CoercionType.Text });
      event.completed();
    }
  }
}

function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
      ? window
      : typeof global !== "undefined"
        ? global
        : undefined;
}

const g = getGlobal() as any;

g.action = action;

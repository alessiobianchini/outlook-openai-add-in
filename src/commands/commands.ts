/* global global, Office, self, window */

import { Configuration, OpenAIApi } from "openai";

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});


function action(event: Office.AddinCommands.Event) {
  var setting = Office.context.roamingSettings.get('openApiToken');
  if (!setting) {
    Office.context.mailbox.item.setSelectedDataAsync("OpenAI token not configured", { coercionType: Office.CoercionType.Text });
    event.completed();
  }
  else {
    if (event.source.id == 'BusinessMail') {
      generateBusinessMail().then(function (selectedText) {
        Office.context.mailbox.item.setSelectedDataAsync(selectedText, { coercionType: Office.CoercionType.Text });
        event.completed();
      });
    }
    else if (event.source.id == 'Tanslate') {
      translateToEnglish().then(function (selectedText) {
        Office.context.mailbox.item.setSelectedDataAsync(selectedText, { coercionType: Office.CoercionType.Text });
        event.completed();
      });
    }
    else if (event.source.id == 'CorrectFormat') {
      correctFormat().then(function (selectedText) {
        Office.context.mailbox.item.setSelectedDataAsync(selectedText, { coercionType: Office.CoercionType.Text });
        event.completed();
      });
    }
  }
}

function generateBusinessMail(): Promise<any> {
  return new Office.Promise(function (resolve, reject) {
    try {
      Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, async function (asyncResult) {
        const configuration = new Configuration({
          apiKey: Office.context.roamingSettings.get('openApiToken'),
        });
        const openai = new OpenAIApi(configuration);
        const response = await openai.createChatCompletion({
          model: "gpt-3.5-turbo",
          messages: [
            {
              role: "system",
              content:
                "You are a helpful assistant that can help users to better manage emails. The following prompt contains the whole mail thread. ",
            },
            {
              role: "user",
              content: `Generate business mail from this text using it's original language: ${asyncResult.value?.data}`,
            },
          ],
        });

        let res = response.data.choices[0].message.content;
        res = res?.replace(/(^"|"$)/g, '');

        resolve(res);
      });
    } catch (error) {
      reject(error);
    }
  });
}

function translateToEnglish(): Promise<any> {
  return new Office.Promise(function (resolve, reject) {
    try {
      Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, async function (asyncResult) {
        const configuration = new Configuration({
          apiKey: Office.context.roamingSettings.get('openApiToken'),
        });
        const openai = new OpenAIApi(configuration);
        const response = await openai.createChatCompletion({
          model: "gpt-3.5-turbo",
          messages: [
            {
              role: "system",
              content:
                "You are a helpful assistant that can help users to better manage emails. The following prompt contains the whole mail thread. ",
            },
            {
              role: "user",
              content: "Translate to english this mail: " + asyncResult.value?.data,
            },
          ],
        });

        let res = response.data.choices[0].message.content;
        res = res?.replace(/(^"|"$)/g, '');

        resolve(res);
      });
    } catch (error) {
      reject(error);
    }
  });
}

function correctFormat(): Promise<any> {
  return new Office.Promise(function (resolve, reject) {
    try {
      Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, async function (asyncResult) {
        const configuration = new Configuration({
          apiKey: Office.context.roamingSettings.get('openApiToken'),
        });
        const openai = new OpenAIApi(configuration);
        const response = await openai.createChatCompletion({
          model: "gpt-3.5-turbo",
          messages: [
            {
              role: "system",
              content:
                "You are a helpful assistant that can help users to better manage emails. The following prompt contains the whole mail thread. ",
            },
            {
              role: "user",
              content: "Correct spelling and grammar: " + asyncResult.value?.data,
            },
          ],
        });

        let res = response.data.choices[0].message.content;
        res = res?.replace(/(^"|"$)/g, '');

        resolve(res);
      });
    } catch (error) {
      reject(error);
    }
  });
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

// The add-in command functions need to be available in global scope
g.action = action;

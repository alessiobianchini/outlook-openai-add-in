/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

import { Configuration, OpenAIApi } from "openai";

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Generate a business mail when the add-in command is executed.
 * @param event
 */
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
          apiKey: "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",
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

        resolve(response.data.choices[0].message.content);
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
          apiKey: "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",
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

        resolve(response.data.choices[0].message.content);
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
          apiKey: "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",
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

        resolve(response.data.choices[0].message.content);
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

function onNewMessageComposeHandler() {
  var setting = Office.context.roamingSettings.get('openApiToken');
  if (!setting) {
    Office.context.ui.displayDialogAsync('https://localhost:3000/tokenpopup.html', { height: 30, width: 20 },
      function (asyncResult) {
        const dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, (args) => {
          Office.context.roamingSettings.set('openApiToken', args);
          dialog.close();
        });
      }
    );
  }
}

const g = getGlobal() as any;

// The add-in command functions need to be available in global scope
g.action = action;
g.onNewMessageComposeHandler = onNewMessageComposeHandler;

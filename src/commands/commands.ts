/* eslint-disable no-undef */
/* eslint-disable prettier/prettier */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */

import OpenAI from "openai";

Office.onReady(() => {
  // If needed, Office.js is ready to be called
});

/**
 * Generate a business mail when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  getSelectedText().then(function (selectedText) {
    Office.context.mailbox.item.setSelectedDataAsync(selectedText, { coercionType: Office.CoercionType.Text });
    event.completed();
  });
}

function getSelectedText(): Promise<any> {
  return new Office.Promise(function (resolve, reject) {
    try {
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, async function (asyncResult) {
        const openai = new OpenAI({ 
          apiKey: process.env.OPENAI_API_KEY,
          dangerouslyAllowBrowser: true
        });

        const response = await openai.chat.completions.create({
        messages: [
          { 
            role: "system", 
            content: "You are a helpful assistant that can help users to better manage emails. The following prompt contains the whole mail thread." 
          },
          { 
            role: "user", 
            content: `Summarize the following mail thread and extract the key points: ${asyncResult.value}` 
          }
        ],
        model: "gpt-4o-mini",
        max_tokens: 150
      });

        resolve(response.choices[0]?.message?.content);
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

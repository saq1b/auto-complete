/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
  }
});

export async function run() {
  return Word.run(async (context) => {
    /**
     * Insert your Word code here
     */
     
    const typedWords = new Set();
    
    function trackWords(context) {
      const body = context.document.body;
      body.getText().then(result => {
        const words = result.value.match(/\b\w+\b/g);
        if (words) {
          words.forEach(word => typedWords.add(word));
        }
      });
    } 

    function getSuggestions(prefix) {
      const lowerPrefix = prefix.toLowerCase();
      const suggestions = Array.from(typedWords).filter(word =>
        word.toLowerCase().startsWith(lowerPrefix)
      );
      return suggestions.slice(0, 5); // Show top 5 suggestions
    }

    function insertText(context, word) {
      const range = context.document.getSelection();
      range.insertText(word, Word.InsertLocation.replace);
    }

    async function getCurrentWord() {
      await Word.run(async (context) => {
        const selection = context.document.getSelection();
        selection.load("text");
        await context.sync();
    
        const text = selection.text;
        const lastWord = text.split(/\s+/).pop();
	console.log(lastWord);
        const suggestions = getSuggestions(lastWord);
        showDropdownInDocument(suggestions);
      });
    }
    Office.context.ui.displayDialogAsync(
      "https://yourdomain.com/suggestions.html",
      { height: 30, width: 20, displayInIframe: true },
      function (asyncResult) {
        // Handle dialog events
      }
    );
    async function insertSuggestion(word) {
      await Word.run(async (context) => {
        const range = context.document.getSelection();
        range.insertText(word + " ", Word.InsertLocation.replace);
        await context.sync();
      });
    }
    // insert a paragraph at the end of the document.
    //const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    //paragraph.font.color = "blue";

    await context.sync();
  });
}

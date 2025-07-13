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

    const inputBox = document.getElementById("textInput");
    
    inputBox.addEventListener("input", (e) => {
      const value = e.target.value;
      const lastWord = value.split(/\s+/).pop();
      
      // Only suggest if it's alphabetic
      if (/^[a-zA-Z]+$/.test(lastWord)) {
        const suggestions = getSuggestions(lastWord);
        showSuggestions(suggestions); // You'll create this function
      }
    
      // Check for space or punctuation: end of word
      if (/\s$/.test(value) || /[.,!?]$/.test(value)) {
        trackWords(); // Update word memory on word completion
      }
    });
    // insert a paragraph at the end of the document.
    //const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // change the paragraph color to blue.
    //paragraph.font.color = "blue";

    await context.sync();
  });
}

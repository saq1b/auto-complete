
// taskpane.js
let typedWords = new Set();

Office.onReady(() => {
  // Add your custom event handlers when Office is ready
  document.getElementById("startTracking").onclick = () => {
    trackWordsFromBody(); // Initial scan
    startMonitoringTyping(); // Begin real-time tracking
  };
});

function trackWordsFromBody() {
  Word.run(async (context) => {
    const body = context.document.body;
    const text = body.text;

    const words = text.match(/\b\w+\b/g);
    if (words) {
      words.forEach((word) => typedWords.add(word));
      console.log("Initial document scan:", Array.from(typedWords));
    }

    await context.sync();
  }).catch((error) => console.error("trackWordsFromBody error:", error));
}

function getSuggestions(prefix) {
  const lowerPrefix = prefix.toLowerCase();
  const suggestions = Array.from(typedWords).filter((word) =>
    word.toLowerCase().startsWith(lowerPrefix)
  );
  return suggestions.slice(0, 5); // Limit results
}

function insertSuggestion(word) {
  Word.run(async (context) => {
    const range = context.document.getSelection();
    range.insertText(word + " ", Word.InsertLocation.replace);
    await context.sync();
  }).catch((error) => console.error("insertSuggestion error:", error));
}

function startMonitoringTyping() {
  setInterval(() => {
    Word.run(async (context) => {
      const selection = context.document.getSelection();
      selection.load("text");
      await context.sync();

      const text = selection.text;
      const lastWord = text.split(/\s+/).pop();

      if (/^[a-zA-Z]+$/.test(lastWord)) {
        const suggestions = getSuggestions(lastWord);
        if (suggestions.length > 0) {
          showDropdown(suggestions);
        }
      }

      if (/\s$/.test(text) || /[.,!?]$/.test(text)) {
        trackWordsFromBody(); // Update memory
      }
    }).catch((error) => console.error("monitorTyping error:", error));
  }, 1000); // Poll every second
}

function showDropdown(suggestions) {
  const container = document.getElementById("suggestionsList");
  container.innerHTML = "";

  suggestions.forEach((suggestion) => {
    const item = document.createElement("div");
    item.className = "suggestion-item";
    item.textContent = suggestion;
    item.onclick = () => {
      insertSuggestion(suggestion);
      container.innerHTML = "";
    };
    container.appendChild(item);
  });
}

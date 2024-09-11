Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    document.getElementById("logStyleContentButton").onclick = getListInfoFromSelection;
  }
});

async function getListInfoFromSelection() {
  try {
    await Word.run(async (context) => {
      console.log("Getting list info and styles from selection");

      const selection = context.document.getSelection();
      const selectionRange = selection.getRange();
      const paragraphs = selectionRange.paragraphs;
      paragraphs.load("items");
      await context.sync();

      console.log(`Total paragraphs in the selection: ${paragraphs.items.length}`);

      let currentList = [];
      let currentLevel = -1;
      let clipboardData = [];
      let parentNumbering = [];

      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        paragraph.load("text,style,isListItem");
        await context.sync();

        const style = paragraph.style;
        const text = paragraph.text.trim();
        const isListItem = paragraph.isListItem;

        if (isListItem) {
          paragraph.listItem.load("level,listString");
          await context.sync();

          const level = paragraph.listItem.level;
          const listString = paragraph.listItem.listString || "";

          if (level <= currentLevel && currentList.length > 0) {
            clipboardData.push(currentList.join("\n"));
            currentList = [];
          }

          if (level <= parentNumbering.length) {
            parentNumbering = parentNumbering.slice(0, level);
          }

          parentNumbering[level] = listString;

          const fullNumbering = parentNumbering.slice(0, level + 1).join(".");
          const indent = "  ".repeat(level);
          const formattedItem = `${fullNumbering} ${text}`;
          currentList.push(`${indent}${formattedItem}`);
          currentLevel = level;
        } else {
          if (currentList.length > 0) {
            clipboardData.push(currentList.join("\n"));
            currentList = [];
            currentLevel = -1;
          }
          clipboardData.push(text);
        }
      }

      if (currentList.length > 0) {
        clipboardData.push(currentList.join("\n"));
      }

      const clipboardString = clipboardData.join("\n");
      copyToClipboard(clipboardString);

      console.log("All data copied to clipboard:");
      console.log(clipboardString);
    });
  } catch (error) {
    console.error("An error occurred:", error);
    if (error instanceof OfficeExtension.Error) {
      console.error("Debug info:", error.debugInfo);
    }
  }
}

function copyToClipboard(text) {
  const textArea = document.createElement("textarea");
  textArea.value = text;

  textArea.style.position = "fixed";
  textArea.style.left = "-999999px";
  textArea.style.top = "-999999px";
  document.body.appendChild(textArea);

  textArea.focus();
  textArea.select();

  try {
    const successful = document.execCommand("copy");
    const msg = successful ? "successful" : "unsuccessful";
    console.log("Copying text was " + msg);
  } catch (err) {
    console.error("Unable to copy to clipboard", err);
  }

  document.body.removeChild(textArea);
}

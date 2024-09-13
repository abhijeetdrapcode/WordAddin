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

      let clipboardData = [];
      let parentNumbering = [];
      let paragraphCounter = 1; // Counter for non-list paragraphs

      for (let i = 0; i < paragraphs.items.length; i++) {
        const paragraph = paragraphs.items[i];
        paragraph.load("text,style,isListItem");
        await context.sync();

        let text = paragraph.text.trim();
        const isListItem = paragraph.isListItem;

        // Remove special non-printable/control characters from text
        text = text.replace(/[^\x20-\x7E]/g, "");

        if (isListItem) {
          paragraph.listItem.load("level,listString");
          await context.sync();

          const level = paragraph.listItem.level;
          const listString = paragraph.listItem.listString || "";

          // Update parentNumbering for the current level
          if (level <= parentNumbering.length) {
            parentNumbering = parentNumbering.slice(0, level);
          }
          parentNumbering[level] = listString;

          const fullNumbering = parentNumbering.slice(0, level + 1).join(".");

          // Format as key-value pair
          clipboardData.push(`"${fullNumbering}": "${text}"`);
        } else {
          // For non-list paragraphs, assign a unique number after "paragraph"
          clipboardData.push(`"paragraph_${paragraphCounter}": "${text}"`);
          paragraphCounter++;
        }
      }

      const clipboardString = `{\n${clipboardData.join(",\n")}\n}`;
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

    // Show the copy success message if copy was successful
    if (successful) {
      const copyMessage = document.getElementById("copyMessage");
      copyMessage.style.display = "block"; // Show the message

      // Hide the message after 15 seconds
      setTimeout(() => {
        copyMessage.style.display = "none";
      }, 15000); // 15 seconds
    }
  } catch (err) {
    console.error("Unable to copy to clipboard", err);
  }

  document.body.removeChild(textArea);
}

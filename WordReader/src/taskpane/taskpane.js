/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office, Word */

Office.onReady((info) => {
  if (info.host === Office.HostType.Word) {
    if (!Office.context.requirements.isSetSupported("WordApi", "1.3")) {
      console.log("Sorry. The add-in uses Word.js APIs that are not available in your version of Office.");
    }

    document.getElementById("insert-paragraph").onclick = insertParagraph;
    document.getElementById("check-first-bold").onclick = checkFirstBold;
    document.getElementById("check-second-underline").onclick = checkSecondUnderline;
    document.getElementById("get-third-size").onclick = getThirdSize;
    document.getElementById("run").onclick = run;

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

export async function run() {
  return Word.run(async (context) => {
    let documentBody = context.document.body;

    console.log(documentBody);

    // Load the text property
    documentBody.load("text");

    // Synchronize the state
    await context.sync();
  });
}
export async function checkFirstBold() {
  Word.run(async (context) => {
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");

    await context.sync();

    if (paragraphs.items.length > 0) {
      // Get the first paragraph
      let firstParagraph = paragraphs.items[0];

      // Load the text of the paragraph
      firstParagraph.load("text");
      await context.sync();

      // Extract the first word
      let words = firstParagraph.text.split(" ");
      if (words.length > 0) {
        let firstWordRange = firstParagraph.search(words[0], { matchCase: true }).getFirst();
        firstWordRange.load("font/bold");

        await context.sync();

        console.log(`First word: "${words[0]}"`);
        console.log(`Is bold? ${firstWordRange.font.bold}`);
      }
    } else {
      console.log("The document is empty.");
    }
  }).catch((error) => {
    console.error("Error: " + error);
  });
}

export async function checkSecondUnderline() {
  Word.run(async (context) => {
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");

    await context.sync();

    if (paragraphs.items.length > 0) {
      // Get the first paragraph
      let firstParagraph = paragraphs.items[0];

      // Load the text of the paragraph
      firstParagraph.load("text");
      await context.sync();

      // Extract words from the paragraph
      let words = firstParagraph.text.split(/\s+/); // Split by spaces
      if (words.length > 1) {
        // Ensure at least two words exist
        let thirdWordRange = firstParagraph.search(words[1], { matchCase: true }).getFirst();
        thirdWordRange.load("font/underline");

        await context.sync();

        console.log(`Second word: "${words[1]}"`);
        console.log(`Is underlined? ${thirdWordRange.font.underline !== "None"}`);
      } else {
        console.log("There is no second word in the paragraph.");
      }
    } else {
      console.log("The document is empty.");
    }
  }).catch((error) => {
    console.error("Error: " + error);
  });
}

export async function getThirdSize() {
  Word.run(async (context) => {
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");

    await context.sync();

    if (paragraphs.items.length > 0) {
      // Get the first paragraph
      let firstParagraph = paragraphs.items[0];

      // Load the text of the paragraph
      firstParagraph.load("text");
      await context.sync();

      // Extract words from the paragraph
      let words = firstParagraph.text.split(/\s+/); // Split by spaces
      if (words.length > 1) {
        // Ensure at least two words exist
        let thirdWordRange = firstParagraph.search(words[2], { matchCase: true }).getFirst();
        thirdWordRange.load("font/size");

        await context.sync();

        console.log(`Third word: "${words[1]}"`);
        console.log(`Size? "${thirdWordRange.font.size}"`);
      } else {
        console.log("There is no third word in the paragraph.");
      }
    } else {
      console.log("The document is empty.");
    }
  }).catch((error) => {
    console.error("Error: " + error);
  });
}

export async function insertParagraph() {
  await Word.run(async (context) => {
    const docBody = context.document.body;
    docBody.insertParagraph("In the small, charming town of Willowbrook, life unfolds at a gentle pace.", "Start");

    await context.sync();
  }).catch(function (error) {
    console.log("Error: " + error);
    if (error instanceof OfficeExtension.Error) {
      console.log("Debug info: " + JSON.stringify(error.debugInfo));
    }
  });
}

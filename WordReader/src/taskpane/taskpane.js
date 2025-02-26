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
    document.getElementById("run-word-task").onclick = runWordTask;
    document.getElementById("check-first-bold").onclick = checkFirstBold;
    document.getElementById("check-second-underline").onclick = checkSecondUnderline;
    document.getElementById("get-third-size").onclick = getThirdSize;
    document.getElementById("load-file").onclick = loadFile;
    document.getElementById("compare-paragraphs").onclick = compareParagraphs;

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
  }
});

function getRandomInt(min, max) {
  return Math.floor(Math.random() * (max - min + 1)) + min;
}

function arrayBufferToBase64(buffer) {
  let binary = "";
  let bytes = new Uint8Array(buffer);
  let len = bytes.byteLength;
  for (let i = 0; i < len; i++) {
    binary += String.fromCharCode(bytes[i]);
  }
  return btoa(binary);
}

export function loadFile() {
  let fileInput = document.getElementById("uploadDocx");
  if (fileInput.files.length === 0) {
    alert("Please select a .docx file first.");
    return;
  }

  let file = fileInput.files[0];
  let reader = new FileReader();

  reader.onload = async function (event) {
    let arrayBuffer = event.target.result;
    let base64 = arrayBufferToBase64(arrayBuffer); // Convert to Base64

    await Word.run(async (context) => {
      let doc = context.document;

      // Insert the file content
      doc.body.insertFileFromBase64(base64, Word.InsertLocation.end);
      
      await context.sync();
      console.log("Word file inserted successfully.");
    }).catch((error) => {
      console.error("Error inserting Word file:", error);
    });
  };

  reader.readAsArrayBuffer(file);
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

async function setOrCancelWordBold(paragraph, n, setOrCancel) {
  paragraph.load("text"); // Load paragraph text
  await paragraph.context.sync();

  let text = paragraph.text;
  let words = text.match(/\S+/g); // Extract words while ignoring spaces

  if (!words || words.length < n) {
    console.log(`There is no ${n}th word in the paragraph.`);
    return;
  }

  // 🔹 Get all word ranges within the paragraph
  let wordRanges = paragraph.getTextRanges([" "], true);
  wordRanges.load("items"); // Load the items before accessing them
  await paragraph.context.sync();

  if (wordRanges.items.length < n) {
    console.log(`Could not locate the ${n}th word.`);
    return;
  }

  let nthWordRange = wordRanges.items[n - 1]; // Get the nth word range
  nthWordRange.load("font"); // Load the font properties
  await paragraph.context.sync();

  nthWordRange.font.bold = setOrCancel; // Apply bold formatting
  await paragraph.context.sync();

  if (setOrCancel) {
    console.log(`Bold applied to the ${n}th word: "${words[n - 1]}"`);
  } else { 
    console.log(`Bold was cancelled on the ${n}th word: "${words[n - 1]}"`);
  }
  
}

async function setOrCancelWordUnderline(paragraph, n, setOrCancel) {
  paragraph.load("text"); // Load paragraph text
  await paragraph.context.sync();

  let text = paragraph.text;
  let words = text.match(/\S+/g); // Extract words while ignoring spaces

  if (!words || words.length < n) {
    console.log(`There is no ${n}th word in the paragraph.`);
    return;
  }

  // 🔹 Get all word ranges within the paragraph
  let wordRanges = paragraph.getTextRanges([" "], true);
  wordRanges.load("items"); // Load the items before accessing them
  await paragraph.context.sync();

  if (wordRanges.items.length < n) {
    console.log(`Could not locate the ${n}th word.`);
    return;
  }

  let nthWordRange = wordRanges.items[n - 1]; // Get the nth word range
  nthWordRange.load("font"); // Load the font properties
  await paragraph.context.sync();

  if (setOrCancel) {
    nthWordRange.font.underline = "Single"; // Apply underline formatting
    await paragraph.context.sync();
    console.log(`Underline applied to the ${n}th word: "${words[n - 1]}"`);
  } else {
    nthWordRange.font.underline = "None"; // Apply underline formatting
    await paragraph.context.sync();
    console.log(`Underline was cancelled on the ${n}th word: "${words[n - 1]}"`);
  }

  
}

async function changeWordSize(paragraph, n) {
  paragraph.load("text"); // Load paragraph text
  await paragraph.context.sync();

  let text = paragraph.text;
  let words = text.match(/\S+/g); // Extract words while ignoring spaces

  if (!words || words.length < n) {
    console.log(`There is no ${n}th word in the paragraph.`);
    return;
  }

  // 🔹 Get all word ranges within the paragraph
  let wordRanges = paragraph.getTextRanges([" "], true);
  wordRanges.load("items"); // Load the items before accessing them
  await paragraph.context.sync();

  if (wordRanges.items.length < n) {
    console.log(`Could not locate the ${n}th word.`);
    return;
  }

  let nthWordRange = wordRanges.items[n - 1]; // Get the nth word range
  nthWordRange.load("font"); // Load the font properties
  await paragraph.context.sync();

  nthWordRange.font.size = 20; // Apply underline formatting
  await paragraph.context.sync();

  console.log(`font size changed on the ${n}th word: "${words[n - 1]}"`);
}

export async function runWordTask() {
  try {
    await Word.run(async (context) => {
      let paragraphs = context.document.body.paragraphs;
      paragraphs.load("items");
      await context.sync();

      if (paragraphs.items.length === 0) {
        console.log("The document is empty.");
        return;
      }

      let firstParagraph = paragraphs.items[0];
      firstParagraph.load("text"); // Load paragraph text
      await context.sync();

      let indexOfBold = 6;
      let indexOfBold1 = 2;
      let indexOfUnderline = 3;
      let indexOfUnderline1 = 4;
      let indexOfChangeSize = 5;

      await setOrCancelWordBold(firstParagraph, indexOfBold, true);
      await setOrCancelWordBold(firstParagraph, indexOfBold1, true);
      await setOrCancelWordUnderline(firstParagraph, indexOfUnderline, true);
      await setOrCancelWordUnderline(firstParagraph, indexOfUnderline1, true);
      await setOrCancelWordBold(firstParagraph, indexOfBold1, false);
      await setOrCancelWordUnderline(firstParagraph, indexOfUnderline1, false);
      await changeWordSize(firstParagraph, indexOfChangeSize);
    });
  } catch (error) {
    console.error("Error: " + error);
  }
}

async function compareParagraphs() {
  await Word.run(async (context) => {
    let paragraphs = context.document.body.paragraphs;
    paragraphs.load("items");
    await context.sync();

    if (paragraphs.items.length < 2) {
      console.log("The document must have at least two paragraphs to compare.");
      return;
    }

    let para1 = paragraphs.items[0]; // First paragraph
    let para2 = paragraphs.items[1]; // Second paragraph

    // 🔹 Extract words from each paragraph
    let words1 = para1.getTextRanges([" "], true);
    let words2 = para2.getTextRanges([" "], true);
    words1.load("items");
    words2.load("items");
    await context.sync();

    let minLength = Math.min(words1.items.length, words2.items.length);

    for (let i = 0; i < minLength; i++) {
      let word1 = words1.items[i];
      let word2 = words2.items[i];

      // 🔹 Load text and font properties for each word
      word1.load(["text", "font/bold", "font/underline", "font/size"]);
      word2.load(["text", "font/bold", "font/underline", "font/size"]);
    }
    await context.sync();


    let isParagraphSame = true;
    // 🔍 Compare each word's properties
    for (let i = 0; i < minLength; i++) {
      let word1 = words1.items[i];
      let word2 = words2.items[i];

      let isTextSame = word1.text === word2.text;
      let isBoldSame = word1.font.bold === word2.font.bold;
      let isUnderlineSame = word1.font.underline === word2.font.underline;
      let isFontSizeSame = word1.font.size === word2.font.size;
      

      if (isTextSame && isBoldSame && isUnderlineSame && isFontSizeSame) {
        isParagraphSame = isParagraphSame && true;
      } else {
        isParagraphSame = isParagraphSame && false;
        console.log(` Word ${i + 1} is different`);
      }
    }

    if (isParagraphSame) {
      console.log("The paragraphs are the same");
    }
  });
}

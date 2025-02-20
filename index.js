$("#run").on("click", () => tryCatch(run));

async function run() {
  await Word.run(async (context) => {
    console.log("Starting run function");
    
    const wordsList = []
    const res = {
      firstWordBold: false,
      secondWordUnderline: '',
      thirdWordFontSize: 0
    }
    console.log("Initialized variables", {wordsList, res});

    // Get all paragraphs first
    console.log("Getting paragraphs");
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();
    console.log("Loaded paragraphs", paragraphs.items.length);

    // Get words from all paragraphs
    for (let i = 0; i < paragraphs.items.length; i++) {
      console.log(`Processing paragraph ${i}`);
      console.log(paragraphs.items[i])
      if(paragraphs.items[i].text === "") {
        continue;
      }
      const words = paragraphs.items[i].split([" "], true /* trimDelimiters*/, true /* trimSpaces */);
      words.load(["text", "font"]);
      await context.sync();
      console.log(`Loaded words from paragraph ${i}`, words.items.length);
      
      words.toJSON().items.forEach(word => {
        wordsList.push(word);
      });
      console.log("Current wordsList length:", wordsList.length);

      if (wordsList.length >= 3) {
        console.log("Found at least 3 words, breaking loop");
        break;
      }
    }

    // Check if we have at least 3 words
    if (wordsList.length < 3) {
      console.error("Not enough words found");
      throw new Error("Document must contain at least 3 words");
    }
    console.log("Found enough words", wordsList);

    // Get font properties of first 3 words
    res.firstWordBold = !!wordsList[0].font.bold;
    res.secondWordUnderline = wordsList[1].font.underline;
    res.thirdWordFontSize = wordsList[2].font.size;
    console.log("Got font properties", res);

    console.log("Final result:", res);
  });
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
  try {
    console.log("Starting try-catch wrapper");
    await callback();
    console.log("Completed successfully");
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error("Error occurred:", error);
  }
}

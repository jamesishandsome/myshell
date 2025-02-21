// $("#run").on("click", () => tryCatch(run));

async function run() {
  await Word.run(async (context) => {

    const wordsList = []
    const res = {
      firstWordBold: false,
      secondWordUnderline: '',
      thirdWordFontSize: 0
    }

    // Get all paragraphs first
    const paragraphs = context.document.body.paragraphs;
    paragraphs.load("text");
    await context.sync();

    // Get words from all paragraphs
    for (let i = 0; i < paragraphs.items.length; i++) {
      if(paragraphs.items[i].text === "") {
        continue;
      }
      const words = paragraphs.items[i].split([" ",",","."], true /* trimDelimiters*/, true /* trimSpaces */);
      words.load(["text", "font"]);
      await context.sync();

      words.toJSON().items.forEach(word => {
        wordsList.push(word);
      });

      if (wordsList.length >= 3) {
        break;
      }
    }

    // Check if there are at least 3 words
    if (wordsList.length < 3) {
      console.error("Not enough words found");
        return;
    }


    // Get font properties of first 3 words
    res.firstWordBold = !!wordsList[0].font.bold;
    res.secondWordUnderline = wordsList[1].font.underline;
    res.thirdWordFontSize = wordsList[2].font.size;
    console.log(res.firstWordBold? "First word is bold" : "First word is not bold");
    console.log(`Second word is ${res.secondWordUnderline} underlined`);
    console.log(`Third word font size is ${res.thirdWordFontSize}`);
  });
}

async function tryCatch(callback) {
  try {
    await callback();
  } catch (error) {
    // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
    console.error("Error occurred:", error);
  }
}



const main = async () => {
    await tryCatch(run);
}

main();
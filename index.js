$("#run").on("click", () => tryCatch(run));

async function run() {
    await Word.run(async (context) => {
        const wordsList = []
        const res  = {
            firstWordBold : false,
            secondWordUnderline : '',
            thirdWordFontSize : 0
        }
        // if the first paragraph contains more than 3 words:
        let paragraph: Word.Paragraph = context.document.body.paragraphs.getFirst();
        const words = paragraph.split([" "], true /* trimDelimiters*/, true /* trimSpaces */);
        words.load(["text","font"]);
        await context.sync();
        words.toJSON().items.map((word) => {
            wordsList.push(word)
        })
        while(wordsList.length<3){
            console.log("get next paragraph")
            paragraph = paragraph.getNext();
            console.log(paragraph)
            const words = paragraph.split([" "], true /* trimDelimiters*/, true /* trimSpaces */);
            words.load(["text","font"]);
            await context.sync();
            words.toJSON().items.map((word) => {
                wordsList.push(word)
            })
        }

        console.log(wordsList)
        // check first word font is bold or not
        res.firstWordBold = !!wordsList[0].font.bold;
        // check second word font is underline or not
        console.log(wordsList[1].font)
        res.secondWordUnderline = wordsList[1].font.underline;
        // check third word font size
        res.thirdWordFontSize = wordsList[2].font.size;

        console.log(res)
    });
}

// Default helper for invoking an action and handling errors.
async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    }
}

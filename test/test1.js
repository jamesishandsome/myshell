import { OfficeMockObject } from "office-addin-mock";
import * as assert from "node:assert";
import run from "../index";

const mockData = {
    context: {
        document: {
            body: {
                paragraph: {
                    font: {},
                    text: "",
                },
                paragraphs: {
                    items:[
                        {
                            text: "Hello World",
                            split: function (delimiters, trimDelimiters, trimSpaces) {
                                return [
                                    {
                                        text: "Hello",
                                        font: {
                                            bold: false,
                                        },
                                    },
                                    {
                                        text: "World",
                                        font: {
                                            underline: "",
                                        },
                                    },
                                    {
                                        text: "",
                                        font: {
                                            size: 0,
                                        },
                                    },
                                ];
                            },
                            load: function (properties) {
                                return;
                            },
                        },
                    ]
                },
                // Mock the Body.insertParagraph method.
                insertParagraph: function (paragraphText, insertLocation) {
                    this.paragraph.text = paragraphText;
                    this.paragraph.insertLocation = insertLocation;
                    return this.paragraph;
                },
            },
        },
    },
    // Mock the Word.InsertLocation enum.
    InsertLocation: {
        end: "end",
    },
    // Mock the Word.run function.
    run: async function(callback) {
        await callback(this.context);
    },
};

// describe(`Run`, function () {
//     it("Word", async function () {
//         const wordMock = new OfficeMockObject(mockData);
//         global.Word = wordMock;
//         const res = await run();
//         assert.strictEqual(res.firstWordBold, false);
//         assert.strictEqual(res.secondWordUnderline, "");
//         assert.strictEqual(res.thirdWordFontSize, 0);
//     });
// });

const wordMock = new OfficeMockObject(mockData);
global.Word = wordMock;
await Word.run(async (context) => {
    // Insert a paragraph at the end of the document.
    const paragraph = context.document.body.insertParagraph("Hello World", Word.InsertLocation.end);

    // Change the font color to blue.
    paragraph.font.color = "blue";

    await context.sync();
});
const res = await run();
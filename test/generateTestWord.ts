import * as path from "node:path";
import {Document, Packer, Paragraph, TextRun} from "docx";


// normal bold single 40
const doc1 = new Document({
    sections: [
        {
            properties: {},
            children: [
                new Paragraph({
                    children: [
                        new TextRun({text: "In ", bold: true}),
                        new TextRun({
                            text: "the ", underline: {
                                type: "single",
                            }
                        }),
                        new TextRun({text: "small", size: 40}),
                        new TextRun({text: " quaint", bold: true}),
                        new TextRun({
                            text: ", charming town of Willowbrook, life unfolds at a gentle pace. The cobblestone streets are lined with colorful, centuries - old houses, each with its own unique story. A meandering river runs through the heart of the town, its waters reflecting the changing hues of the sky.\n" +
                                " \n" +
                                "On sunny days, locals gather in the central square. Elderly men play chess under the shade of ancient oak trees, while children chase each other, their laughter filling the air. The local bakery, with its warm, freshly baked bread aroma, is always a popular spot. Inside, the bakers work tirelessly, creating delicious pastries and loaves that are the pride of the town.\n" +
                                " \n" +
                                "As the evening approaches, the soft glow of streetlights illuminates the town. Couples take leisurely walks along the riverbank, enjoying the peaceful scenery. Willowbrook is not just a place; it's a community bound by shared traditions, a place where time seems to slow down, and where the simple joys of life are cherished every day.\n",
                        }),
                    ],
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc1).then(async (buffer) => {
    await Bun.write(path.join(__dirname, "test1.docx"), buffer);
});


// mixed bold and mixed underline and mixed size
const doc2 = new Document({
    sections: [
        {
            properties: {},
            children: [
                new Paragraph({
                    children: [
                        new TextRun({text: "I", bold: true}),
                        new TextRun({text: "n ", bold: false}),
                        new TextRun({
                            text: "th", underline: {
                                type: "none",
                            }
                        }),
                        new TextRun({
                            text: "e ", underline: {
                                type: "single",
                            }
                        }),
                        new TextRun({text: "sma", size: 100}),
                        new TextRun({text: "ll", size: 20}),
                        new TextRun({text: " quaint", bold: true}),
                        new TextRun({
                            text: ", charming town of Willowbrook, life unfolds at a gentle pace. The cobblestone streets are lined with colorful, centuries - old houses, each with its own unique story. A meandering river runs through the heart of the town, its waters reflecting the changing hues of the sky.\n" +
                                " \n" +
                                "On sunny days, locals gather in the central square. Elderly men play chess under the shade of ancient oak trees, while children chase each other, their laughter filling the air. The local bakery, with its warm, freshly baked bread aroma, is always a popular spot. Inside, the bakers work tirelessly, creating delicious pastries and loaves that are the pride of the town.\n" +
                                " \n" +
                                "As the evening approaches, the soft glow of streetlights illuminates the town. Couples take leisurely walks along the riverbank, enjoying the peaceful scenery. Willowbrook is not just a place; it's a community bound by shared traditions, a place where time seems to slow down, and where the simple joys of life are cherished every day.\n",
                        }),
                    ],
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc2).then(async (buffer) => {
    await Bun.write(path.join(__dirname, "test2.docx"), buffer);
});


// Start from empty paragraph + period or full stop + bold + double + 100
const doc3 = new Document({
    sections: [
        {
            properties: {},
            children: [
                new Paragraph({
                    children:[
                        new TextRun({text: " "}),
                    ]
                }),
                new Paragraph({
                    children: [
                        new TextRun({text: "On ", bold: true}),
                        new TextRun({
                            text: "sunny ", underline: {
                                type: "double",
                            }
                        }),
                        new TextRun({text: "days", size: 100}),
                        new TextRun({
                            text:
                                ", locals gather in the central square. Elderly men play chess under the shade of ancient oak trees, while children chase each other, their laughter filling the air. The local bakery, with its warm, freshly baked bread aroma, is always a popular spot. Inside, the bakers work tirelessly, creating delicious pastries and loaves that are the pride of the town.\n" +
                                " \n" +
                                "As the evening approaches, the soft glow of streetlights illuminates the town. Couples take leisurely walks along the riverbank, enjoying the peaceful scenery. Willowbrook is not just a place; it's a community bound by shared traditions, a place where time seems to slow down, and where the simple joys of life are cherished every day.\n",
                        }),
                    ],
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc3).then(async (buffer) => {
    await Bun.write(path.join(__dirname, "test3.docx"), buffer);
});


// start from empty paragraph + comma + bold + double + 100
const doc4  = new Document({
    sections: [
        {
            properties: {},
            children: [
                new Paragraph({
                    children:[
                        new TextRun({text: " "}),
                    ]
                }),
                new Paragraph({
                    children: [
                        new TextRun({text: ", ", bold: false}),
                        new TextRun({text: "locals ", bold: true}),
                        new TextRun({
                            text: "gather ", underline: {
                                type: "double",
                            }
                        }),
                        new TextRun({text: "in ", size: 100}),
                        new TextRun({
                            text:
                                "the central square. Elderly men play chess under the shade of ancient oak trees, while children chase each other, their laughter filling the air. The local bakery, with its warm, freshly baked bread aroma, is always a popular spot. Inside, the bakers work tirelessly, creating delicious pastries and loaves that are the pride of the town.\n" +
                                " \n" +
                                "As the evening approaches, the soft glow of streetlights illuminates the town. Couples take leisurely walks along the riverbank, enjoying the peaceful scenery. Willowbrook is not just a place; it's a community bound by shared traditions, a place where time seems to slow down, and where the simple joys of life are cherished every day.\n",
                        }),
                    ],
                }),
            ],
        },
    ],
});

Packer.toBuffer(doc4).then(async (buffer) => {
    await Bun.write(path.join(__dirname, "test4.docx"), buffer);
});


// total 2 words
const doc5 = new Document({
    sections: [
        {
            properties: {},
            children: [
                new Paragraph({
                    children: [
                        new TextRun({text: "In ", bold: true}),
                        new TextRun({
                            text: "the ", underline: {
                                type: "single",
                            }
                        }),
                        ],
                }),
            ],
        },
    ],
});


Packer.toBuffer(doc5).then(async (buffer) => {
    await Bun.write(path.join(__dirname, "test5.docx"), buffer);
});
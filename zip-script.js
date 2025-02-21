import compressing from 'compressing';
import convert from 'xml-js';
import WordExtractor from 'word-extractor';
import * as fs from "node:fs";

export const run = async (file) => {
    const word = await new WordExtractor().extract(`${file}.docx`)
    const wordString = await word.getBody()

// delete all , and . and split by space
    const firstThreeWords = wordString.trim().replace(/,/g, '').replace(/\./g, '').split(' ').filter((word) => {
        return word !== ''
    }).slice(0, 3)
    if (firstThreeWords.length < 3) {
        console.log('Not enough words found')
        throw new Error('Not enough words found')
    }

    await compressing.zip.uncompress(`${file}.docx`, file)

    const document = Bun.file(`${file}/word/document.xml`)

    const result = convert.xml2json(await document.text())
    const json = JSON.parse(result)
    const body = json.elements[0].elements[0]

    const wordsList = []
    try {
        body.elements.map((p) => {
            p.elements.filter((element) => {
                return element.type === 'element'
            }).map((element) => {
                    if (element.name === 'w:r') {
                        wordsList.push(
                            {
                                bold: (() => {
                                    for (const subElement of element.elements) {
                                        if (subElement.name === 'w:rPr') {
                                            for (const rPrElement of subElement.elements) {
                                                if (rPrElement.name === 'w:b' && !rPrElement.attributes) {
                                                    return true
                                                }
                                                if (rPrElement.name === 'w:b' && rPrElement.attributes && rPrElement.attributes['w:val']) {
                                                    return rPrElement.attributes['w:val'] !== 'false'
                                                }
                                            }
                                        }
                                    }
                                    return false;
                                })(),
                                underline: (() => {
                                    for (const subElement of element.elements) {
                                        if (subElement.name === 'w:rPr') {
                                            for (const rPrElement of subElement.elements) {
                                                if (rPrElement.name === 'w:u' && rPrElement.attributes && rPrElement.attributes['w:val']) {
                                                    return rPrElement.attributes['w:val']
                                                }
                                            }
                                        }
                                    }
                                    return "none"
                                })(),
                                fontSize: element.elements.filter((element) => {
                                    return element.name === 'w:rPr'
                                }).filter((element) => {
                                    return element.elements.filter((element) => {
                                        return element.name === 'w:sz'
                                    }).length > 0
                                }).map((element) => {
                                    return element.elements.filter((element) => {
                                        return element.name === 'w:sz'
                                    })[0].attributes["w:val"]
                                })[0],
                                text: (() => {
                                    for (const subElement of element.elements) {
                                        if (subElement.name === 'w:t') {
                                            if (!subElement.elements) {
                                                return ' '
                                            }
                                            return subElement.elements[0].text
                                        }
                                    }
                                    return ""
                                })()
                            }
                        )
                    } else {
                        console.log()
                    }
                }
            )
        })
        // handle first word:
        const firstWord = []
        const firstWordBold = []
        const secondWord = []
        const secondWordUnderline = []
        const thirdWord = []
        const thirdWordFontSize = []


        wordsList.map((element, index) => {
            element.text.split('').map((text, index) => {
                if (firstWord.join('') === firstThreeWords[0] && secondWord.join('') === firstThreeWords[1] && thirdWord.join('') === firstThreeWords[2]) {

                } else if (text === " " || text === "," || text === ".") {
                } else if (firstWord.join('') !== firstThreeWords[0]) {
                    firstWord.push(text)
                    if (element.bold) {
                        firstWordBold.push(true)
                    } else {
                        firstWordBold.push(false)
                    }
                } else if (secondWord.join('') !== firstThreeWords[1]) {
                    secondWord.push(text)
                    // if element is [] ,return false. if it is object[], return JSON.stringfy(element)
                    if (element.underline.length === 0) {
                        secondWordUnderline.push(false)
                    }
                    if (element.underline.length > 0) {
                        secondWordUnderline.push(element.underline)
                    }
                } else if (thirdWord.join('') !== firstThreeWords[2]) {
                    thirdWord.push(text)
                    thirdWordFontSize.push(element.fontSize)
                }
            })
        })


        // if all in firstWordBold are true, then firstWordBold is true, if all false or undefined, then firstWordBold is false. if some are true, some are false, then firstWordBold is mixed
        const firstWordBoldResult = firstWordBold.every((val, i, arr) => val === arr[0]) ? firstWordBold[0] : 'mixed'
        const secondWordUnderlineResult = secondWordUnderline.every((val, i, arr) => val === arr[0]) ? secondWordUnderline[0] : 'mixed'
        const thirdWordFontSizeResult = thirdWordFontSize.every((val, i, arr) => val === arr[0]) ? thirdWordFontSize[0] : 'mixed'

        console.log(firstWordBoldResult)
        console.log(secondWordUnderlineResult)
        console.log(thirdWordFontSizeResult)

        // delete the extracted folder
        fs.rmdirSync(`${file}`, {recursive: true})

        return {
            firstWordBold: firstWordBoldResult,
            secondWordUnderline: secondWordUnderlineResult,
            thirdWordFontSize: thirdWordFontSizeResult
        }


    } catch (err) {
        console.log(err)
    }
}


await run('./test/test1')
const fs = require('fs');
const pdf = require('pdf-parse');
const {
    Document,
    Packer,
    Paragraph,
} = require('docx');

const getTextToArr = (text) => (text.split('\n\n'));

const getDocxToText = async (dataBuffer) => {
    return pdf(dataBuffer)
    .then(function(result) {
        return result.text; // docx to text
    })
    .catch(function(reason) {
        console.log('.doc or .docx file has some problems');
        console.error(`reason : ${reason}`);    
    });
}

const ChangeTextToDocxFile = async (textArr, filePath, fileName) => {
    console.log("Now converting file ...");
    let sectionsArr = [];

    for(let i = 1; i < textArr.length; i++) {
        sectionsArr.push({
            children: [
                new Paragraph({
                    text: textArr[i]
                })
            ]
        })
    }

    const docx = new Document({
        sections: sectionsArr
    });

    Packer.toBuffer(docx).then((buffer) => {
        let docxFileName = fileName.split('.pdf')[0];
        fs.writeFileSync(`${filePath}/${docxFileName}.docx`, buffer);
    });

}

const convertDocx = (filePath, fileName) => {
    
    let dataBuffer = fs.readFileSync(filePath + '/' + fileName);

    getDocxToText(dataBuffer)
    .then(function(result) { // result : text
        return getTextToArr(result);
    })
    .catch(function(reason) {
        console.error(`reason : ${reason}`);    
    })
    .then(function(textArr) {
        ChangeTextToDocxFile(textArr, filePath, fileName);
    });
}

exports.convertDocx = convertDocx;
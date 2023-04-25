const docx = require("docx");
const { Document, Packer, Paragraph, TextRun, ImageRun } = docx;

module.exports = genDocx = async (data) => {
  const image = new ImageRun({
    data: data.image,
    transformation: {
      width: 600,
      height: 400,
    },
  });
  const doc = new Document({
    sections: [
      {
        properties: {},
        children: [
          new Paragraph({
            children: [image],
          }),
          new Paragraph({
            children: [
              new TextRun({
                text: `Le Rapport  :\n`,
                size: 30,
                underline: true,
                color: "42f58a",
              }),
              new TextRun({
                text: `On note que le pourcentage de ${data["first"]} est de ${data["firstPer"]}, ce qui est supérieur au pourcentage de ${data["second"]} qui possèdent  ${data["secondPer"]}`,
                size: 25,
              }),
            ],
          }),
        ],
      },
    ],
  });
  const b64string = await Packer.toBase64String(doc);
  return b64string;
};

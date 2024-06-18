import { Document, 
  Header, 
  Footer, 
  Table, 
  TableRow, 
  TableCell, 
  Paragraph, 
  Packer, 
  VerticalAlign, 
  WidthType, 
  AlignmentType, 
  BorderStyle, 
  TextRun, 
  PageNumber } from 'docx';
import { saveAs } from "file-saver"
import { useState, createRef } from "react";

const DatasheetTemplate = (props) => {
    const colorTemperature = props.data["Correlated Colour Temperature (CCT) in K"];
    const ra = props.data["Ra"];
    const mmCode = props.code;
    const [productCode, setProductCode] = useState('');
    const productCodeRef = createRef();
    const unit = props.logistic["unit"];
    const boxLength = props.logistic["length"];
    const boxWidth = props.logistic["width"];
    const boxHeight = props.logistic["height"];
    const weight = props.logistic["weight"];

    const font = "Arial"
    const contentFont = "Arial";
    const contentFontSize = 16;

    const currentDate = new Date().toLocaleDateString();

    const headerBorders = { 
      top: { style: BorderStyle.NIL, size: 0,},
      bottom: { style: BorderStyle.NIL, size: 0,},
      left: { style: BorderStyle.NIL, size: 0,},
      right: { style: BorderStyle.NIL, size: 0,},
    };

    const tableBorders = { 
      top: { style: BorderStyle.NIL, size: 0,},
      bottom: { style: BorderStyle.SINGLE, size: 0.75, color: "BFBFBF"},
      left: { style: BorderStyle.NIL, size: 0,},
      right: { style: BorderStyle.NIL, size: 0,},
    };

    const headerTable = new Table({
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Series name"],
                      font: font,
                      size: 56,
                      color: "00B0F0"
                    })
                  ],
                }),
              ],
              borders: headerBorders,
              width: {
                size: 5000,
                type: WidthType.DXA,
              },
            }),
            new TableCell({
              children: [
              ],
              borders: headerBorders,
              width: {
                size: 5000,
                type: WidthType.DXA,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [ new Paragraph ({
                  children: [
                    new TextRun ({
                      text: props.data["Product Type"],
                      font: font,
                      size: 24,
                    }),
                  ]
                })
              ],
              borders: headerBorders,
              verticalAlign: VerticalAlign.CENTER,
              width: {
                size: 50,
                type: WidthType.PERCENTAGE,
              },
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "PRODUCT DATASHEET",
                      font: font,
                      size: 24,
                    })
                  ],
                  alignment: AlignmentType.RIGHT, 
                }),
              ],
              borders: headerBorders,
              width: {
                size: 50,
                type: WidthType.PERCENTAGE,
              },
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun ({
                      text: props.data["Customer Model No.                (NEW ErP)"],
                      font: font,
                      size: 20,
                    }),
                  ]
                }),
              ],
              borders: headerBorders,
              verticalAlign: VerticalAlign.CENTER,
              width: {
                size: 50,
                type: WidthType.PERCENTAGE,
              },
            }),
            new TableCell({
              children: [],
              borders: headerBorders,
              verticalAlign: VerticalAlign.CENTER,
              width: {
                size: 50,
                type: WidthType.PERCENTAGE,
              },
            }),
          ],
        }),
      ],
    });

    const footerTable = new Table({
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "www.megaman.com",
                      size: 12,
                    })
                  ] 
                }),
              ],
              verticalAlign: VerticalAlign.CENTER,
              width: {
                size: 33,
                type: WidthType.PERCENTAGE,
              },
              borders: headerBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      children: [PageNumber.CURRENT],
                      size: 12,
                    }),
                    new TextRun({
                        children: [" of ", PageNumber.TOTAL_PAGES],
                        size: 12,
                    }),
                  ],

                  alignment: AlignmentType.CENTER
                }),
              ],
              verticalAlign: VerticalAlign.CENTER,
              width: {
                size: 34,
                type: WidthType.PERCENTAGE,
              },
              borders: headerBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: currentDate,
                      size: 12,
                    }),
                  ],
                  alignment: AlignmentType.RIGHT
                }),
              ],
              verticalAlign: VerticalAlign.CENTER,
              width: {
                size: 33,
                type: WidthType.PERCENTAGE,
              },
              borders: headerBorders,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "info@megaman.cc",
                      size: 12,
                    }),
                  ],
                  alignment: AlignmentType.LEFT
                }),
              ],
              verticalAlign: VerticalAlign.CENTER,
              width: {
                size: 33,
                type: WidthType.PERCENTAGE,
              },
              borders: headerBorders,
            }),
            new TableCell({
              children: [],
              verticalAlign: VerticalAlign.CENTER,
              width: {
                size: 34,
                type: WidthType.PERCENTAGE,
              },
              borders: headerBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Version 1.0",
                      size: 12,
                    }),
                  ],
                  alignment: AlignmentType.RIGHT,
                }),
              ],
              verticalAlign: VerticalAlign.CENTER,
              width: {
                size: 33,
                type: WidthType.PERCENTAGE,
              },
              borders: headerBorders,
            }),
          ],
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "© Copyright 2024. All rights reserved by MEGAMAN®",
                      size: 12,
                    }),
                  ],
                }),
              ],
              verticalAlign: VerticalAlign.CENTER,
              width: {
                size: 33,
                type: WidthType.PERCENTAGE,
              },
              borders: headerBorders,
            }),
            new TableCell({
              children: [],
              verticalAlign: VerticalAlign.CENTER,
              width: {
                size: 34,
                type: WidthType.PERCENTAGE,
              },
              borders: headerBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Data subject to change",
                      size: 12,
                    }),
                  ],
                  alignment: AlignmentType.RIGHT,
                }),
              ],
              verticalAlign: VerticalAlign.CENTER,
              width: {
                size: 33,
                type: WidthType.PERCENTAGE,
              },
              borders: headerBorders,
              fontSize: 8,
            }),
          ],
        }),
      ],
    });

    const colourTempCode = () => {
      let code;

      if (ra >= 80 && ra < 90) {
          code = '8';
      } else if (ra >= 90) {
          code = '9';
      }
      console.log(code);

      if (typeof colorTemperature === 'string') {
          const temps = colorTemperature.replace(/\//g, ',').split(',');
          let index = 0;
          temps.forEach((temp) => {
              index++;
              code += temp.replace(/\D+/g, '').substr(0, 2);
              if (index === 1) {
                  if (ra >= 80 && ra < 90) {
                      code += '8';
                  } else if (ra >= 90) {
                      code += '9';
                  }
              }
          });
          console.log(code);
      } else {
          code += colorTemperature/100;
          console.log(code);
      }

      setProductCode(code);
      productCodeRef.current = code
      console.log(productCode);
    };

    const productInfo = new Paragraph({
      children: [
        new TextRun({
          text: "PRODUCT INFORMATION",
          font: contentFont,
          size: contentFontSize,
          color: "00B0F0",
        }),
      ],
      spacing: { after: 100, before: 100,},
    });

    const productInfoTable = new Table({
      borders: tableBorders,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Model Number",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Customer Model No.                (NEW ErP)"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Product Code",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Customer Model No.                (NEW ErP)"]+"+"+productCodeRef.current+"+"+props.data["Fitting Color"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "MM Code",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: mmCode,
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
      ],
    });

    const electricalInfo = new Paragraph({
      children: [
        new TextRun({
          text: "ELECTRICAL INFOMATION",
          font: contentFont,
          size: contentFontSize,
          color: "00B0F0",
        }),
      ],
      spacing: { after: 100, before: 100,},
    });

    const electricalInfoTable = new Table({
      borders: tableBorders,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Input Voltage",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Rated Voltage       (V)"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Frequency",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Frequency(Hz)"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Input Current",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Input Current     (mA)"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Power",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["On-mode power  (Pon)  （W）"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Power Factor",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Power factor"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Total harmonic distortion",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["THD"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Surge Protection",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Surge Voltage(L-N)  [V]"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Inrush Current",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Inrush current Ipeak    (A) "],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Inrush Duration",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Inrush current Twidth (uS) "],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
      ],
    });

    const dimming = new Paragraph({
      children: [
        new TextRun({
          text: "DIMMING AND CONTROLS",
          font: contentFont,
          size: contentFontSize,
          color: "00B0F0",
        }),
      ],
      spacing: { after: 100, before: 100,},
    });

    const dimmabilityTable = new Table({
      borders: tableBorders,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Dimmability",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Dimmable"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
      ],
    });

    const photometricalInfo = new Paragraph({
      children: [
        new TextRun({
          text: "PHOTOMETRICAL INFOMATION",
          font: contentFont,
          size: contentFontSize,
          color: "00B0F0",
        }),
      ],
      spacing: { after: 100, before: 100,},
    });

    const photometricalInfoTable = new Table({
      borders: tableBorders,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Luminous Flux",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Total luminous flux              (lm)"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Luminous Efficacy",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Total mains efficacy ηΤM (lm/W)"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Colour Temperature",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Correlated Colour Temperature (CCT) in K"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "CRI",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Ra"]+"",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Beam Angle",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Beam angle    (°)"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "UGR",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["UGR          (if Applicable) "],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Colour Consistency",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Colour consistency [MacAdam ellips steps]"]+"",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Max. intensity",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Maximum intensity  (cd)"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Flickering",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Flickering                   ( %)"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "SVM",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Stroboscopic effect metric (SVM)"]+"",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Pst LM",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Flicker metric   (Pst LM)"]+"",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
      ],
    });

    const lifePerformance = new Paragraph({
      children: [
        new TextRun({
          text: "LIFE PERFORMANCE",
          font: contentFont,
          size: contentFontSize,
          color: "00B0F0",
        }),
      ],
      spacing: { after: 100, before: 100,},
    });

    const lifePerformanceTable = new Table({
      borders: tableBorders,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Lifetime",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Norminal Life (h) \r\nas L**B** lifetime"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Number of Switching Cycles",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Switching Cycles"]+"",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
      ],
    });

    const standardsInfo = new Paragraph({
      children: [
        new TextRun({
          text: "STANDARDS AND APPLICATION",
          font: contentFont,
          size: contentFontSize,
          color: "00B0F0",
        }),
      ],
      spacing: { after: 100, before: 100,},
    });

    const standardsInfoTable = new Table({
      borders: tableBorders,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Protection Class",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Protection Class"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Glow Wire",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Glow-wire (℃）"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Photobiological Safety Group",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Photobiological Risk Group"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Energy Class",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Energy efficiency class"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Standards Compliance",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["SCG Standards Compliance"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
      ],
    });

    const installInfo = new Paragraph({
      children: [
        new TextRun({
          text: "INSTALLATION AND CAPABILITIES",
          font: contentFont,
          size: contentFontSize,
          color: "00B0F0",
        }),
      ],
      spacing: { after: 100, before: 100,},
    });

    const installInfoTable = new Table({
      borders: tableBorders,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Installation",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Mounting"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "B10",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["MCB:B10（pcs)"] + "",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "B16",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["MCB:B16（pcs)"] + "",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "C10",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["MCB:C10（pcs)"] + "",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "C16",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["MCB:C16（pcs)"] + "",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
      ],
    });

    const mechanicalInfo = new Paragraph({
      children: [
        new TextRun({
          text: "MECHANICAL AND METERIAL",
          font: contentFont,
          size: contentFontSize,
          color: "00B0F0",
        }),
      ],
      spacing: { after: 100, before: 100,},
    });

    const mechanicalInfoTable = new Table({
      borders: tableBorders,
      width: {
        size: 100,
        type: WidthType.PERCENTAGE,
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Optical Material",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Diffuser Material"],
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Housing Material",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Housing Material"] + "",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Housing Colour",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Fitting Color"] + "",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Diameter",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Diameter\r\n(mm)"] + "",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Width",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Width (mm)"] + "",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Length",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Length (mm)"] + "",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Height",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Height (mm)"] + "",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Cut-out",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Recessed Cut Out (mm)"] + "",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Weights",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 40,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: props.data["Net Weight (g)"] + "",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 60,
                type: WidthType.PERCENTAGE,
              },
              borders: tableBorders,
            }),
          ]
        }),
      ],
    });

    const wiringDiagram = new Paragraph({
      children: [
        new TextRun({
          text: "WIRING DIAGRAM",
          font: contentFont,
          size: contentFontSize,
          color: "00B0F0",
        }),
      ],
      spacing: { after: 100, before: 100,},
    });

    const photometricDiagram = new Paragraph({
      children: [
        new TextRun({
          text: "PHOTOMETRIC DIAGRAM",
          font: contentFont,
          size: contentFontSize,
          color: "00B0F0",
        }),
      ],
      spacing: { after: 100, before: 100,},
    });

    const logisticInfo = new Paragraph({
      children: [
        new TextRun({
          text: "LOGISTIC INFORMATION",
          font: contentFont,
          size: contentFontSize,
          color: "00B0F0",
        }),
      ],
      spacing: { after: 100, before: 100,},
    });

    const logisticInfoTable = new Table({
      width: {
        size: 9638,
        type: WidthType.DXA,
      },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "MM code",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 2000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
              shading: {
                fill: "00B0F0",
                color: "auto",
              },
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Packaging Unit",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  alignment: AlignmentType.CENTER, 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 1638,
                type: WidthType.DXA,
              },
              borders: headerBorders,
              shading: {
                fill: "00B0F0",
                color: "auto",
              },
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Outer Box Dimensions",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  alignment: AlignmentType.CENTER, 
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              columnSpan: 3,
              width: {
                size: 3000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
              shading: {
                fill: "00B0F0",
                color: "auto",
              },
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "Gross Weight per Outer Box",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ], 
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 3000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
              shading: {
                fill: "00B0F0",
                color: "auto",
              },
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 2000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
              shading: {
                fill: "00B0F0",
                color: "auto",
              },
            }),
            new TableCell({
              children: [],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 1638,
                type: WidthType.DXA,
              },
              borders: headerBorders,
              shading: {
                fill: "00B0F0",
                color: "auto",
              },
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "L",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 20, before: 20,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 1000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
              shading: {
                fill: "00B0F0",
                color: "auto",
              },
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "W",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 20, before: 20,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 1000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
              shading: {
                fill: "00B0F0",
                color: "auto",
              },
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "H",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 20, before: 20,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 1000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
              shading: {
                fill: "00B0F0",
                color: "auto",
              },
            }),
            new TableCell({
              children: [],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 3000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
              shading: {
                fill: "00B0F0",
                color: "auto",
              },
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 2000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
              shading: {
                fill: "00B0F0",
                color: "auto",
              },
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "pcs/unit",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 100, before: 20,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 1638,
                type: WidthType.DXA,
              },
              borders: headerBorders,
              shading: {
                fill: "00B0F0",
                color: "auto",
              },
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "mm",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 100, before: 20,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 1000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
              shading: {
                fill: "00B0F0",
                color: "auto",
              },
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "mm",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 100, before: 20,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 1000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
              shading: {
                fill: "00B0F0",
                color: "auto",
              },
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "mm",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 100, before: 20,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 1000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
              shading: {
                fill: "00B0F0",
                color: "auto",
              },
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: "kg",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 100, before: 20,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 3000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
              shading: {
                fill: "00B0F0",
                color: "auto",
              },
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: mmCode,
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 2000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: unit+"",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 1638,
                type: WidthType.DXA,
              },
              borders: headerBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: boxLength+"",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 1000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: boxWidth+"",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 1000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: boxHeight+"",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 1000,
                type: WidthType.DXA,
              },
              borders: headerBorders,
            }),
            new TableCell({
              children: [
                new Paragraph ({
                  children: [
                    new TextRun({
                      text: weight+"",
                      size: contentFontSize,
                      font: contentFont,
                    }),
                  ],
                  alignment: AlignmentType.CENTER,
                  spacing: { after: 100, before: 100,},
                }),
              ],
              verticalAlign: VerticalAlign.LEFT,
              width: {
                size: 3000,
                type: WidthType.DXA,
              },
              borders: headerBorders
            }),
          ]
        }),
      ],
    });

    const handleCreateDocx = async () => {
        const doc = new Document({
          sections: [
            {
              properties:{
                page: {
                  margin: {
                    top: 1000, // 1 inch
                    right: 1000, // 1 inch
                    bottom: 1000, // 1 inch
                    left: 1000, // 1 inch
                  }
                }
              },
              headers: {
                default: new Header({
                  children: [headerTable],
                }),
              },
              footers: {
                default: new Footer({
                  children: [footerTable],
                }),
              },
              children: [
                productInfo,
                productInfoTable,
                electricalInfo,
                electricalInfoTable,
                dimming,
                dimmabilityTable,
                photometricalInfo,
                photometricalInfoTable,
                lifePerformance,
                lifePerformanceTable,
                standardsInfo,
                standardsInfoTable,
                installInfo,
                installInfoTable,
                mechanicalInfo,
                mechanicalInfoTable,
                wiringDiagram,
                photometricDiagram,
                logisticInfo,
                logisticInfoTable,
              ],
              
            },
          ],
        });
      
        Packer.toBlob(doc).then(blob => {
          console.log(blob);
          saveAs(blob, "example.docx");
          console.log("Document created successfully");
        });
      };

    

    return (
      <button
        onClick={handleCreateDocx}
        className="bg-blue-500 hover:bg-blue-600 text-white font-medium py-2 px-4 rounded-full focus:outline-none focus:ring-2 focus:ring-blue-500 focus:ring-opacity-50"
        >
        Create Datasheet
      </button>
    );
}

export default DatasheetTemplate;
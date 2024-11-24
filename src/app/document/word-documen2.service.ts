import { Injectable } from '@angular/core';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun, WidthType, AlignmentType, HeadingLevel, BorderStyle, ShadingType, TableLayoutType, VerticalAlign } from 'docx';
import { saveAs } from 'file-saver';

@Injectable({
  providedIn: 'root'
})
export class DocumentService2 {
  generateDocument(): void {
    const defaultParagraphStyle = {
      spacing: { before: 120, after: 120, line: 360 },
      style: { font: { size: 24 } }
    };

    const defaultCellMargins = {
      top: 100,
      bottom: 100,
      left: 120,
      right: 120
    };

    const doc = new Document({
      sections: [{
        properties: {},
        children: [
          new Paragraph({
            text: "ANEXO N° 03",
            alignment: AlignmentType.CENTER,
            heading: HeadingLevel.HEADING_1,
            spacing: { after: 200 }
          }),

          // Tabla separada para "Orden de verificación"
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1 },
              bottom: { style: BorderStyle.SINGLE, size: 1 },
              left: { style: BorderStyle.SINGLE, size: 1 },
              right: { style: BorderStyle.SINGLE, size: 1 },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    margins: defaultCellMargins,
                    children: [
                      new Paragraph({
                        alignment: AlignmentType.CENTER,
                        spacing: { before: 100, after: 100 },
                        children: [
                          new TextRun({
                            text: "Acta de Verificación",
                            size: 28,
                            bold: true,
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              })
            ]
          }),

          new Paragraph({
            text: "",
            spacing: { before: 100, after: 50 }
          }),

          // Tabla principal con el resto del contenido
          new Table({
            width: { size: 100, type: WidthType.PERCENTAGE },
            borders: {
              top: { style: BorderStyle.SINGLE, size: 1 },
              bottom: { style: BorderStyle.SINGLE, size: 1 },
              left: { style: BorderStyle.SINGLE, size: 1 },
              right: { style: BorderStyle.SINGLE, size: 1 },
            },
            rows: [
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    spacing: { before: 100, after: 100 },
                                    children: [
                                      new TextRun({
                                        text: "Acta de Verificación N° (CODIGO DE LA OSPE)-2024-VCA-(NUMERO DEL CASO)-026-001",
                                        bold: true,
                                        size: 21,
                                      }),
                                    ],
                                  }),
                                ],
                                columnSpan: 2,
                              }),
                            ],
                          }),
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({
                                        text: "En ............... , siendo las ........... horas del día .... del mes ............ del año ............, los verificadores de EsSalud que suscriben la presente Acta, nos constituimos en el local de la Entidad Empleadora .................................., sito en ........................................ con el fin de realizar la Verificación de (colocar nombres y apellidos del asegurado a verificar), identificado con DNI .................., iniciada mediante Orden de Verificación N° ................. por la modalidad de ......................, constatándose lo siguiente:",
                                      }),
                                    ],
                                  }),
                                ],
                                columnSpan: 2,
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),
              // Presente y texto explicativo
              new TableRow({
                children: [
                  new TableCell({
                    margins: defaultCellMargins,
                    children: [
                        new Paragraph({
                            spacing: { before: 10, after: 10 },
                            children: [
                            new TextRun({ text: "I. Participantes de la verificación",
                                bold: true,
                                size: 21,
                             }),
                            ],
                        }),
                    ]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    margins: defaultCellMargins,
                    children: [
                      new Paragraph({
                        text: "Personal y representantes de la entidad empleadora.",
                        spacing: { before: 10, after: 10 }
                      }),
                    ]
                  })
                ]
              }),
              // Tabla de asegurados
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                            new TableRow({
                              children: [
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [
                                    new Paragraph({
                                      alignment: AlignmentType.CENTER, // Centrar el texto
                                      children: [
                                        new TextRun({
                                          text: "N°",
                                          bold: true, // Negrita
                                        }),
                                      ],
                                    }),
                                  ],
                                  width: { size: 10, type: WidthType.PERCENTAGE },
                                }),
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [
                                    new Paragraph({
                                      alignment: AlignmentType.CENTER, // Centrar el texto
                                      children: [
                                        new TextRun({
                                          text: "Nombres y Apellidos",
                                          bold: true, // Negrita
                                        }),
                                      ],
                                    }),
                                  ],
                                  width: { size: 45, type: WidthType.PERCENTAGE },
                                }),
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [
                                    new Paragraph({
                                      alignment: AlignmentType.CENTER, // Centrar el texto
                                      children: [
                                        new TextRun({
                                          text: "DNI",
                                          bold: true, // Negrita
                                        }),
                                      ],
                                    }),
                                  ],
                                  width: { size: 15, type: WidthType.PERCENTAGE },
                                }),
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [
                                    new Paragraph({
                                      alignment: AlignmentType.CENTER, // Centrar el texto
                                      children: [
                                        new TextRun({
                                          text: "Cargo",
                                          bold: true, // Negrita
                                        }),
                                      ],
                                    }),
                                  ],
                                  width: { size: 30, type: WidthType.PERCENTAGE },
                                }),
                              ],
                            }),
                            ...Array(3)
                              .fill(null)
                              .map(() =>
                                new TableRow({
                                  children: [
                                    new TableCell({margins: defaultCellMargins, children: [new Paragraph("")] }),
                                    new TableCell({margins: defaultCellMargins, children: [new Paragraph("")] }),
                                    new TableCell({margins: defaultCellMargins, children: [new Paragraph("")] }),
                                    new TableCell({margins: defaultCellMargins, children: [new Paragraph("")] }),
                                  ],
                                })
                              ),
                          ]
                      })
                    ]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    margins: defaultCellMargins,
                    children: [
                      new Paragraph({
                        text: "Verificadores designados por EsSalud:",
                        spacing: { before: 10, after: 10 }
                      }),
                    ]
                  })
                ]
              }),
              // Texto verificadores y tabla
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    alignment: AlignmentType.CENTER, // Centrar el texto
                                    children: [
                                      new TextRun({
                                        text: "N°",
                                        bold: true, // Negrita
                                      }),
                                    ],
                                  }),
                                ],
                                width: { size: 10, type: WidthType.PERCENTAGE },
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [
                                      new TextRun({
                                        text: "Apellidos y Nombres",
                                        bold: true,
                                      }),
                                    ],
                                  }),
                                ],
                                width: { size: 90, type: WidthType.PERCENTAGE },
                              }),
                            ],
                          }),
                          ...Array.from({ length: 2 }, () =>
                            new TableRow({
                              children: [
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [new Paragraph({ text: "" })],
                                }),
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [new Paragraph({ text: "" })],
                                }),
                              ],
                            })
                          ),
                        ],
                      }),
                    ],
                  }),
                ],
              }),

              new TableRow({
                children: [
                  new TableCell({
                    margins: defaultCellMargins,
                    children: [
                        new Paragraph({
                            spacing: { before: 10, after: 10 },
                            children: [
                            new TextRun({ text: "II. Generalidades",
                                bold: true,
                                size: 21,
                             }),
                            ],
                        }),
                    ]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    margins: defaultCellMargins,
                    children: [
                      new Paragraph({
                        text: "Dirección de la entidad:",
                        spacing: { before: 10, after: 10 }
                      }),
                    ]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({
                                        text: "Telefono°",
                                      }),
                                    ],
                                  }),
                                ],
                                width: { size: 30, type: WidthType.PERCENTAGE },
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({
                                        text: "Horarios de trabajo:",
                                      }),
                                    ],
                                  }),
                                ],
                                width: { size: 60, type: WidthType.PERCENTAGE },
                              }),
                            ],
                          }),
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({
                                        text: "N° de Asegurados:",
                                      }),
                                    ],
                                  }),
                                ],
                                width: { size: 30, type: WidthType.PERCENTAGE },
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({
                                        text: "Tiene Sucursales: Sí (    )   No (    )  De ser afirmativa la respuesta, anote su Dirección:",
                                      }),
                                    ],
                                  }),
                                ],
                                width: { size: 60, type: WidthType.PERCENTAGE },
                              }),
                            ],
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),

              new TableRow({
                children: [
                  new TableCell({
                    margins: defaultCellMargins,
                    children: [
                        new Paragraph({
                            spacing: { before: 10, after: 10 },
                            children: [
                            new TextRun({ text: "Sucursal 1:",
                             }),
                            ],
                        }),
                    ]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    margins: defaultCellMargins,
                    children: [
                        new Paragraph({
                            spacing: { before: 10, after: 10 },
                            children: [
                            new TextRun({ text: "Sucursal 2:",
                             }),
                            ],
                        }),
                    ]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    margins: defaultCellMargins,
                    children: [
                        new Paragraph({
                            spacing: { before: 10, after: 10 },
                            children: [
                            new TextRun({ text: "¿Se encuentra en el Centro de Trabajo el asegurado sujeto de verificación?  Sí (    )   No (    )",
                             }),
                            ],
                        }),
                    ]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    margins: defaultCellMargins,
                    children: [
                        new Paragraph({
                            spacing: { before: 10, after: 10 },
                            children: [
                            new TextRun({ text: "Si no se encuentra, indicar el motivo de su ausencia:",
                             }),
                            ],
                        }),
                    ]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    margins: defaultCellMargins,
                    children: [
                        new Paragraph({
                            spacing: { before: 10, after: 10 },
                            children: [
                            new TextRun({ text: "¿Le presenta los documentos que sustentan la ausencia del asegurado? Sí (    )   No (    ).............................................",
                             }),
                            ],
                        }),
                    ]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    margins: defaultCellMargins,
                    children: [
                        new Paragraph({
                            spacing: { before: 10, after: 10 },
                            children: [
                            new TextRun({ text: "III. Requerimientos de Información (Estos deben exhibirse y presentarse en copias legibles, (colocar Sí o No)",
                                bold: true,
                                size: 21,
                             }),
                            ],
                        }),
                    ]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Requerimiento de Información y/o Documentos", bold: true }),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              }),
                              new TableCell({
                                width: { size: 4, type: WidthType.PERCENTAGE },
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "¿Cumplió con presentar?", bold: true }),
                                    ],
                                  }),
                                ],
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Observación", bold: true }),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              })
                            ]
                          }),
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Contrato de Trabajo del asegurado" }),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              }),
                              new TableCell({
                                width: { size: 4, type: WidthType.PERCENTAGE },
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "", bold: true }),
                                    ],
                                  }),
                                ],
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Plazo Indeterminado: "}),
                                      new TextRun({ text: "Sí (  ) No (  )" }),
                                      new TextRun({ text: "Tipo de contrato:" }),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              })
                            ]
                          }),
                          //3
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Declaraciones Tributarias PDT- últimos seis (6) meses - PDT 621 de los periodos a verificar"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              }),
                              new TableCell({
                                width: { size: 4, type: WidthType.PERCENTAGE },
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ""}),
                                    ],
                                  }),
                                ],
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ":"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              })
                            ]
                          }),
                          //4
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Registros especiales según su actividad económica"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              }),
                              new TableCell({
                                width: { size: 4, type: WidthType.PERCENTAGE },
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ""}),
                                    ],
                                  }),
                                ],
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Registrado en el MTPE: Sí (    )   No (    )"}),
                                      new TextRun({ text: "Otro registro:"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              })
                            ]
                          }),
                          //5
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Planilla de Sueldos o Remuneraciones-PDT 601"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              }),
                              new TableCell({
                                width: { size: 4, type: WidthType.PERCENTAGE },
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ""}),
                                    ],
                                  }),
                                ],
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Pago Gratificación Julio: Sí ( ) No ( )  S/."}),
                                      new TextRun({ text: "Pago Gratificación Diciembre: Sí ( ) No (    )  S/. "}),
                                      new TextRun({ text: "Pago CTS Noviembre: Sí ( ) No ( )  S/. "}),
                                      new TextRun({ text: "Pago CTS Mayo: Sí ( ) No ( ) S/. "}),
                                      new TextRun({ text: "Pago Vacaciones: Sí ( ) No ( ) S/. "}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              })
                            ]
                          }),
                          //6
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Boletas de pago del asegurado de los últimos seis (6) meses de los periodos a verificar"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              }),
                              new TableCell({
                                width: { size: 4, type: WidthType.PERCENTAGE },
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ""}),
                                    ],
                                  }),
                                ],
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ":"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              })
                            ]
                          }),
                          //7
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Partes diarios de asistencia de los últimos seis (6) meses de los periodos a verificar"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              }),
                              new TableCell({
                                width: { size: 4, type: WidthType.PERCENTAGE },
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ""}),
                                    ],
                                  }),
                                ],
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ":"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              })
                            ]
                          }),
                          //8
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Seis (6) últimos pagos de Aporte ONP/AFP del asegurado"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              }),
                              new TableCell({
                                width: { size: 4, type: WidthType.PERCENTAGE },
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ""}),
                                    ],
                                  }),
                                ],
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ":"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              })
                            ]
                          }),
                          //9
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Descansos médicos de los últimos seis (6) meses"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              }),
                              new TableCell({
                                width: { size: 4, type: WidthType.PERCENTAGE },
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ""}),
                                    ],
                                  }),
                                ],
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ":"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              })
                            ]
                          }),
                          //10
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Documentos de asignación de funciones del asegurado"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              }),
                              new TableCell({
                                width: { size: 4, type: WidthType.PERCENTAGE },
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ""}),
                                    ],
                                  }),
                                ],
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ":"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              })
                            ]
                          }),
                          //11
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Documentos presentados por el trabajador al empleador que evidencien las funciones desarrolladas en los últimos seis (6) ) meses"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              }),
                              new TableCell({
                                width: { size: 4, type: WidthType.PERCENTAGE },
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ""}),
                                    ],
                                  }),
                                ],
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ":"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              })
                            ]
                          }),
                          //12
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Registro de Ventas - Comprobantes de pago que emite"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              }),
                              new TableCell({
                                width: { size: 4, type: WidthType.PERCENTAGE },
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ""}),
                                    ],
                                  }),
                                ],
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ":"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              })
                            ]
                          }),
                          //13
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Registro de Compras"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              }),
                              new TableCell({
                                width: { size: 4, type: WidthType.PERCENTAGE },
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ""}),
                                    ],
                                  }),
                                ],
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ":"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              })
                            ]
                          }),
                          //14
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: "Relación de productos o servicios que vende"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              }),
                              new TableCell({
                                width: { size: 4, type: WidthType.PERCENTAGE },
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ""}),
                                    ],
                                  }),
                                ],
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph({
                                    children: [
                                      new TextRun({ text: ":"}),
                                    ],
                                  }),
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              })
                            ]
                          }),
                        ]
                      }),

                    ]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    margins: defaultCellMargins,
                    children: [
                        new Paragraph({
                            spacing: { before: 10, after: 10 },
                            children: [
                            new TextRun({ text: "IV. Descripción del lugar donde desarrolla la labor",
                                bold: true,
                                size: 21,
                             }),
                            ],
                        }),
                        new Paragraph({
                          spacing: { before: 10, after: 10 },
                          children: [
                          new TextRun({ text: "Sea por separado y con firma del empleador y/o asegurado",
                           }),
                          ],
                        }),
                        new Paragraph({
                          spacing: { before: 10, after: 10 },
                          children: [
                          new TextRun({ text: "........................................................................................................................",
                           }),
                          ],
                        }),
                    ]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    margins: defaultCellMargins,
                    children: [
                        new Paragraph({
                            spacing: { before: 10, after: 10 },
                            children: [
                            new TextRun({ text: "V. Declaración del empleador sobre la relación jurídica con el asegurado sujeto de verificación",
                                bold: true,
                                size: 21,
                             }),
                            ],
                        }),
                        new Paragraph({
                          spacing: { before: 10, after: 10 },
                          children: [
                          new TextRun({ text: "5.1 ¿Cuál es la relación familiar o de parentesco que tiene con el asegurado?"}),
                          ],
                        }),
                        new Paragraph({
                          spacing: { before: 10, after: 10 },
                          children: [
                            new TextRun({
                              text: "...........................................................................................................................",
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: { before: 10, after: 10 },
                          children: [
                            new TextRun({
                              text: "5.2. ¿Desde cuándo labora el asegurado en su empresa, cuál es su horario y su remuneración?",
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: { before: 10, after: 10 },
                          children: [
                            new TextRun({
                              text: "...........................................................................................................................",
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: { before: 10, after: 10 },
                          children: [
                            new TextRun({
                              text: "5.3. ¿Cuáles son las funciones asignadas al asegurado y especifique en dónde las realiza?",
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: { before: 10, after: 10 },
                          children: [
                            new TextRun({
                              text: "...........................................................................................................................",
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: { before: 10, after: 10 },
                          children: [
                            new TextRun({
                              text: "5.4. ¿Conoce si su trabajador realiza otras actividades? ¿Dónde? ¿En qué horario?",
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: { before: 10, after: 10 },
                          children: [
                            new TextRun({
                              text: "...........................................................................................................................",
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: { before: 10, after: 10 },
                          children: [
                            new TextRun({
                              text: "5.5. Acepto que este documento y las demás comunicaciones me sean remitidas vía correo electrónico",
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: { before: 5, after: 5 },
                          children: [
                            new TextRun({
                              text: "SI [   ]",
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: { before: 5, after: 5 },
                          children: [
                            new TextRun({
                              text: "  Correo electrónico: _______________________",
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: { before: 10, after: 10 },
                          children: [
                            new TextRun({
                              text: "5.6. Acepto la toma de fotos y videos en relación a mi caso de verificación",
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: { before: 5, after: 5 },
                          children: [
                            new TextRun({
                              text: "SI [   ]",
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: { before: 10, after: 10 },
                          children: [
                            new TextRun({
                              text: "5.7. Otros datos adicionales:",
                            }),
                          ],
                        }),
                        new Paragraph({
                          spacing: { before: 10, after: 10 },
                          children: [
                            new TextRun({
                              text: "...........................................................................................................................",
                            }),
                          ],
                        }),
                    ]
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    margins: defaultCellMargins,
                    children: [
                      new Paragraph({
                        spacing: { before: 10, after: 10 },
                        children: [
                          new TextRun({
                            text: "VI. Observaciones: (EN ESTE ACAPITE COLOCAR TODAS LAS AFIRMACIONES QUE EL VERIFICADOR CONSIDERE IMPORTANTES, EN EL DESARROLLO DEL PROCESO. COLOCAR, DE SER EL CASO, LA RESPUESTA O MANIFESTACIÓN DEL EMPLEADOR RESPECTO A LA VERIFICACIÓN REALIZADA).",
                            bold: true,
                            size: 21,
                          }),
                        ],
                      }),
                      new Paragraph({
                        spacing: { before: 10, after: 10 },
                        children: [
                          new TextRun({
                            text: "...........................................................................................................................",
                          }),
                        ],
                      }),
                      new Paragraph({
                        spacing: { before: 10, after: 50 },
                        children: [
                          new TextRun({
                            text: "Siendo las.................. horas del día .............................., se da por concluida la visita de Verificación y se firma la presente acta, teniéndose por notificada en el momento. Dentro de los diez (10) días hábiles siguientes a la fecha consignada en la presente acta, la entidad empleadora podrá presentar los descargos y medios probatorios que estime pertinente y sustenten las declaraciones y/o actuaciones efectuadas.",
                          }),
                        ],
                      }),
                      // Tabla para las firmas con celdas invisibles
                      new Table({
                        rows: [
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: {
                                  top: 0,
                                  bottom: 0,
                                  left: 0,
                                  right: 0,
                                },
                                borders: {
                                  top: { size: 0, style: BorderStyle.NONE },
                                  bottom: { size: 0, style: BorderStyle.NONE },
                                  left: { size: 0, style: BorderStyle.NONE },
                                  right: { size: 0, style: BorderStyle.NONE },
                                },
                                children: [
                                  new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [
                                      new TextRun({
                                        text: "_____________________________",
                                        underline: {},
                                      }),
                                    ],
                                  }),
                                  new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [
                                      new TextRun({
                                        text: "Firma y sello del verificador de EsSalud",
                                        bold: true,
                                      }),
                                    ],
                                  }),
                                ],
                              }),
                              new TableCell({
                                margins: {
                                  top: 0,
                                  bottom: 0,
                                  left: 0,
                                  right: 0,
                                },
                                borders: {
                                  top: { size: 0, style: BorderStyle.NONE },
                                  bottom: { size: 0, style: BorderStyle.NONE },
                                  left: { size: 0, style: BorderStyle.NONE },
                                  right: { size: 0, style: BorderStyle.NONE },
                                },
                                children: [
                                  new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [
                                      new TextRun({
                                        text: "_____________________________",
                                        underline: {},
                                      }),
                                    ],
                                  }),
                                  new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [
                                      new TextRun({
                                        text: "Firma y Sello Entidad Empleadora",
                                        bold: true,
                                      }),
                                    ],
                                  }),
                                  new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [
                                      new TextRun({
                                        text: "Nombre: ",
                                      }),
                                    ],
                                  }),
                                  new Paragraph({
                                    alignment: AlignmentType.CENTER,
                                    children: [
                                      new TextRun({
                                        text: "DNI: ",
                                      }),
                                    ],
                                  }),
                                ],
                              }),
                            ],
                          }),
                        ],
                        width: {
                          size: 100,
                          type: WidthType.PERCENTAGE,
                        },
                        borders: {
                          top: { size: 0, style: BorderStyle.NONE },
                          bottom: { size: 0, style: BorderStyle.NONE },
                          left: { size: 0, style: BorderStyle.NONE },
                          right: { size: 0, style: BorderStyle.NONE },
                        },
                        layout: TableLayoutType.FIXED, // Asegura un diseño fijo sin fluctuaciones
                      }),


                      new Paragraph({
                        spacing: { before: 10, after: 10 },
                        children: [
                          new TextRun({
                            text: "Nota: Forman parte de la presente acta las declaraciones del asegurado y de terceros participantes.",
                          }),
                        ],
                      }),
                      new Paragraph({
                        spacing: { before: 10, after: 10 },
                        children: [
                          new TextRun({
                            text: "La negativa de suscripción de la presente Acta, no invalida el acto de verificación, de conformidad con el artículo 19° del D.S. N° 002-2009-TR.",
                          }),
                        ],
                      }),
                    ],
                  }),
                ],
              }),






            ]
          })
        ]
      }]
    });

    // Generar y descargar el documento
    Packer.toBlob(doc).then(blob => {
      saveAs(blob, "orden_verificacion03.docx");
    });
  }
}

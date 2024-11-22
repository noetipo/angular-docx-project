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
                        text: "Personal y representantes de la entidad empladora.",
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
                                    new TableCell({ children: [new Paragraph("")] }),
                                    new TableCell({ children: [new Paragraph("")] }),
                                    new TableCell({ children: [new Paragraph("")] }),
                                    new TableCell({ children: [new Paragraph("")] }),
                                  ],
                                })
                              ),
                          ]
                      })
                    ]
                  })
                ]
              }),
              // Texto verificadores y tabla
              new TableRow({
                children: [
                  new TableCell({
                    children: [
                      new Paragraph({
                        text: "Verificadores designados por EsSalud:",
                        spacing: { before: 10, after: 10 }
                      }),
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
                                  width: { size: 60, type: WidthType.PERCENTAGE },
                                }),
                              ],
                            }),
                            ...Array(2)
                              .fill(null)
                              .map(
                                () =>
                                  new TableRow({
                                    children: [
                                      new TableCell({ children: [new Paragraph("")] }),
                                      new TableCell({ children: [new Paragraph("")] }),
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
                    children: [
                      new Paragraph({
                        text: "Dirección de la entidad:",
                        spacing: { before: 10, after: 10 }
                      }),
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
                                        new Paragraph({
                                            alignment: AlignmentType.CENTER, // Centrar el texto
                                            children: [
                                              new TextRun({
                                                text: "Telefono°",
                                              }),
                                            ],
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
                                                text: "Horarios de trabajo:",
                                            }),
                                        ],
                                    }),
                                  ],
                                  width: { size: 60, type: WidthType.PERCENTAGE },
                                }),
                              ],
                            }),
                            ...Array(2)
                              .fill(null)
                              .map(
                                () =>
                                  new TableRow({
                                    children: [
                                      new TableCell({ children: [new Paragraph("")] }),
                                      new TableCell({ children: [new Paragraph("")] }),
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
                    children: [
                      new Paragraph({
                        text: "En virtud de lo dispuesto en el Artículo 11° del Reglamento aprobado por Decreto Supremo N° 002-2009-TR, le solicitamos poner a disposición del personal verificador los siguientes documentos:",
                        spacing: { before: 200, after: 200 }
                      }),
                      new Table({
                        width: { size: 100, type: WidthType.PERCENTAGE },
                        rows: [
                          new TableRow({
                            children: [
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph("1) Contrato de Trabajo del asegurado"),
                                  new Paragraph("2) Declaraciones Tributarias-PDT 621- últimos seis (6) meses"),
                                  new Paragraph("3) Registro en el MTPE"),
                                  new Paragraph("4) Registros especiales según su actividad económica"),
                                  new Paragraph("5) Planilla de Sueldos o Remuneraciones/Planilla Electrónica-PDT 601"),
                                  new Paragraph("6) Boletas de Pago del asegurado de los últimos seis (6) meses"),
                                  new Paragraph("7) Partes diarios de Asistencia de los últimos seis (6) meses"),
                                  new Paragraph("8) Seis (6) últimos pagos de Aporte ONP/AFP del asegurado"),
                                  new Paragraph("9) Descansos médicos de los últimos seis (6) meses"),
                                  new Paragraph("10) Documentos de asignación de funciones del asegurado")
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              }),
                              new TableCell({
                                width: { size: 4, type: WidthType.PERCENTAGE },
                                margins: defaultCellMargins,
                                children: [new Paragraph("")]
                              }),
                              new TableCell({
                                margins: defaultCellMargins,
                                children: [
                                  new Paragraph("11) Documentos presentados por el trabajador al empleador que evidencien las funciones desarrolladas en los últimos seis (6) meses"),
                                  new Paragraph("12) Registro de Ventas"),
                                  new Paragraph("13) Comprobantes de pago que emite"),
                                  new Paragraph("14) Registro de Compras"),
                                  new Paragraph("15) Relación de Productos o Servicios que vende"),
                                  new Paragraph("16) Carta de empleador para depósito de CTS de los últimos semestres"),
                                  new Paragraph("17) Título de propiedad, Título de Posesión, Contrato de Arrendamiento o Cesión de Uso del terreno donde se realiza la Actividad Agraria"),
                                  new Paragraph("18) Puede considerar otros documentos"),
                                  new Paragraph("19)	Otros: . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . Colocar el número de ítems solicitados: . . ")
                                ],
                                width: { size: 48, type: WidthType.PERCENTAGE }
                              })
                            ]
                          })
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
                        text: "El verificador, de acuerdo a lo dispuesto en el artículo 07° del Reglamento aprobado por Decreto Supremo N° 002-2009-TR, está facultado para iniciar la verificación inmediatamente después de recibida la Orden de Verificación, ingresar al centro de trabajo, levantar actas, practicar cualquier diligencia de investigación, examen o prueba que considere necesario, requerir información e identificación de las personas que se encuentren en el centro de trabajo materia de la acción de verificación y solicitar la comparecencia de la entidad empleadora o sus representantes, de los trabajadores y de cualesquiera sujetos incluidos en su ámbito de actuación en el centro inspeccionado. El empleador debe permitir el ingreso a los funcionarios y/o servidores públicos en el centro de trabajo, lugar o establecimiento donde se lleva a cabo la verificación, colaborar con ellos durante su visita y facilitar la información y documentación que le sea solicitada para desarrollar la función de verificación, el incumplimiento de lo señalado en el párrafo anterior constituye infracción tipificada en el artículo 25° del Reglamento en mención, estando sujetos a las sanciones contenidas en el anexo de Tabla de Infracciones y Sanciones contenidas en el referido Decreto Supremo. En caso el verificador no sea atendido o exista demora en la atención, llevará a cabo una nueva visita dentro de los tres (3) días hábiles siguientes, para lo cual se deberá tener toda la información y/o documentación a su disposición, tal como se señala en el artículo 16° de la norma antes acotada. Si usted desea confirmar la identidad de los servidores designados podrá acceder a la siguiente dirección electrónica http://www.essalud.gob.pe/agencias-y-oficinas-de-seguros/  y/o comunicarse telefónicamente al número (Teléfono y anexo de la UCF u OSPE, según el tipo de oficina) en el (horario de atención) para comprobar su identidad. Base Legal: Ley N° 29135 reglamentada por el Decreto Supremo N° 002-2009-TR.",
                        spacing: { before: 200, after: 200 }
                      })
                    ],
                    width: { size: 50, type: WidthType.PERCENTAGE }
                  })
                ]
              }),
              new TableRow({
                children: [
                  new TableCell({
                    margins: defaultCellMargins,
                    children: [
                      new Paragraph({
                        text: "Acepto que este documento y las demás comunicaciones me sean remitidas vía correo electrónico   SI    ",
                        spacing: { before: 200, after: 100 }
                      }),
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: "Correo:",
                            bold: true
                          })
                        ],
                        spacing: { before: 100, after: 100 }
                      }),
                      new Paragraph({
                        text: "Acepto la toma de fotos y videos en relación a mi caso de verificación.   SI    ",
                        spacing: { before: 100, after: 100 }
                      }),
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: "_____________________________",
                            underline: {}
                          })
                        ],
                        spacing: { before: 200, after: 50 }
                      }),
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: "Firma y Sello del Jefe de OSPE",
                            bold: true
                          })
                        ],
                        alignment: AlignmentType.LEFT,
                        spacing: { before: 50, after: 100 }
                      }),
                      new Paragraph({
                        children: [
                          new TextRun({
                            text: "Recepción: ............",
                            bold: true
                          })
                        ],
                        alignment: AlignmentType.RIGHT,
                        spacing: { before: 50, after: 100 }
                      })
                    ],
                    width: { size: 50, type: WidthType.PERCENTAGE }
                  })
                ]
              })

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

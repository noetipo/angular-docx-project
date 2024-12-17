import { Injectable } from '@angular/core';
import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  Header,
  Footer,
  ImageRun,
  TabStopType,
  TabStopPosition, BorderStyle, Table, TableRow, TableCell, WidthType
} from 'docx';
import { saveAs } from 'file-saver';
import { logoEssalud } from './logo';
import { imagenPie } from './imegenPie';
@Injectable({
  providedIn: 'root'
})
export class DocumentEssalud2Service {
  generateDocument2(): void {
    const defaultParagraphSpacing = {spacing: {before: 200, after: 200}};

    const base64Data = logoEssalud.split(',')[1];
    const imageBytes = Uint8Array.from(atob(base64Data), (c) => c.charCodeAt(0));

    const base64DataPie = imagenPie.split(',')[1];
    const imageBytesPie = Uint8Array.from(atob(base64DataPie), (c) => c.charCodeAt(0));

    const imagePath = "./assets/imagen.jpg";
    const doc = new Document({
      numbering: {
        config: [
          {
            reference: "numbering1",
            levels: [
              {
                level: 0,
                format: "decimal",
                text: "%1.",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: {left: 0, hanging: 400},
                  },
                  run: {
                    font: "Arial",
                    bold: true,
                  },
                },
              },
              {
                level: 1,
                format: "decimal",
                text: "%1.%2",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: {left: 1200, hanging: 200},
                  },
                  run: {
                    font: "Arial",
                    bold: true,
                  },
                },
              },
            ],
          },
          {
            reference: "numbering2",
            levels: [
              {
                level: 0,
                format: "upperRoman",
                text: "%1.",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: {left: 800, hanging: 400},
                  },
                  run: {
                    font: "Arial",
                    bold: true,
                  },
                },
              },
            ],
          },
          {
            reference: "numbering3",
            levels: [
              {
                level: 0,
                format: "decimal",
                text: "%1.",
                alignment: AlignmentType.LEFT,
                start: 2,
                style: {
                  paragraph: {
                    indent: {left: 0, hanging: 400},
                  },
                  run: {
                    font: "Arial",
                    bold: true,
                  },
                },
              },
              {
                level: 1,
                format: "decimal",
                text: "%1.%2",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: {left: 1200, hanging: 200},
                  },
                  run: {
                    font: "Arial",
                    bold: true,
                  },
                },
              },
            ],
          },
          {
            reference: "numbering4",
            levels: [
              {
                level: 0,
                format: "decimal",
                text: "%1.",
                alignment: AlignmentType.LEFT,
                start: 3,
                style: {
                  paragraph: {
                    indent: {left: 0, hanging: 400},
                  },
                  run: {
                    font: "Arial",
                    bold: true,
                  },
                },
              },
              {
                level: 1,
                format: "decimal",
                text: "%1.%2",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: {left: 1200, hanging: 200},
                  },
                  run: {
                    font: "Arial",
                    bold: true,
                  },
                },
              },
            ],
          },
          {
            reference: "numbering5",
            levels: [
              {
                level: 0,
                format: "decimal",
                text: "%1.",
                alignment: AlignmentType.LEFT,
                start: 4,
                style: {
                  paragraph: {
                    indent: {left: 0, hanging: 400},
                  },
                  run: {
                    font: "Arial",
                    bold: true,
                  },
                },
              },
              {
                level: 1,
                format: "decimal",
                text: "%1.%2",
                alignment: AlignmentType.LEFT,
                style: {
                  paragraph: {
                    indent: {left: 1200, hanging: 200},
                  },
                  run: {
                    font: "Arial",
                    bold: true,
                  },
                },
              },
            ],
          },
        ],
      },
      sections: [
        {
          properties: {},
          headers: {
            default: new Header({
              children: [
                new Table({
                  rows: [
                    new TableRow({
                      children: [
                        new TableCell({
                          borders: { // Bordes invisibles
                            top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                            bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                            left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                            right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                          },
                          children: [
                            new Paragraph({
                              children: [
                                new ImageRun({
                                  data: imageBytes,
                                  transformation: { width: 100, height: 32 },
                                  type: 'png',
                                }),
                              ],
                            }),
                          ],
                          verticalAlign: 'center',
                          width: { size: 20, type: WidthType.PERCENTAGE }, // Ancho de la celda
                        }),
                        new TableCell({
                          borders: { // Bordes invisibles
                            top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                            bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                            left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                            right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                          },
                          children: [
                            new Paragraph({
                              alignment: AlignmentType.CENTER,
                              children: [
                                new TextRun({
                                  text: "Año del Bicentenario, de la consolidación de nuestra Independencia, y de la conmemoración de las heroicas batallas de Junín y Ayacucho",
                                  font: 'Arial',
                                  size: 16,
                                  color: '606060', // Color más suave
                                }),
                              ],
                            }),
                          ],
                          verticalAlign: 'center',
                          width: { size: 80, type: WidthType.PERCENTAGE }, // Ancho de la celda
                        }),
                      ],
                    }),
                  ],
                  width: { size: 100, type: WidthType.PERCENTAGE },
                  borders: { // Bordes de la tabla invisibles
                    top: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    bottom: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    left: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                    right: { style: BorderStyle.NONE, size: 0, color: "FFFFFF" },
                  },
                }),
              ],
            }),
          },


          footers: {
            default: new Footer({
              children: [
                new Paragraph({
                  alignment: AlignmentType.LEFT,
                  children: [
                    new TextRun({
                      text:"Esta es una copia autenticada imprimible de un documento electrónico archivado por ESSALUD, aplicando lo dispuesto por el Art. 25 del D.S. 070-2013- PCM y la Tercera Disposición Complementaria Final del D.S. 026-2016-PCM.",
                      font: "Arial",
                      size: 14,
                    }),
                  ],
                  spacing: { after: 250 },
                }),
                new Paragraph({
                  children: [
                    new ImageRun({
                      data: imageBytesPie,
                      transformation: { width: 600, height: 40 },
                      type: 'png',
                    }),
                  ],
                  alignment: AlignmentType.LEFT,
                }),
              ],
            }),
          },
          children: [
            new Paragraph({
              alignment: AlignmentType.JUSTIFIED,
              children: [
                new TextRun({
                  text: "NOTA N°",
                  font: "Arial",
                  bold: true,
                  size: 20,
                }),
                new TextRun({
                  text: "\t       -OSPE\t\t        -GCSPE-ESSALUD-202X",
                  font: "Arial",
                  size: 20,
                  bold: true,
                }),
              ],
            }),
            new Paragraph({
              spacing: { before: 200, after: 200 },
              children: [
                new TextRun({
                  text: "Lima,",
                  font: "Arial",
                  size: 20,
                }),
              ],
            }),
            // Saludo
            new Paragraph({
              children: [
                new TextRun({
                  text: "Señor(a)",
                  font: "Arial",
                  size: 20,
                }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "NNNN AAPP AAMM",
                  font: "Arial",
                  bold: true,
                  size: 20,
                }),
                new TextRun({
                  text: " (colocar nombre del director o gerente, según corresponda)",
                  font: "Arial",
                  size: 20,
                }),
              ],
            }),

            // Cargos
            new Paragraph({
              children: [
                new TextRun({
                  text: "Director de la IPRESS o",
                  font: "Arial",
                  size: 20,
                  bold: true,
                }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "Gerente de la Red Prestacional / Asistencial u ",
                  font: "Arial",
                  size: 20,
                  bold: true,
                }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "Órgano Prestador Nacional ",
                  font: "Arial",
                  size: 20,
                  bold: true,
                }),
                new TextRun({
                  text: "(colocar el cargo, según corresponda)",
                  font: "Arial",
                  size: 20,
                }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "Dirección",
                  font: "Arial",
                  size: 20,
                }),
              ],
              spacing: { after: 80 },

            }),

            // Asunto
            new Paragraph({
              children: [
                new TextRun({
                  text: "Asunto: ",
                  font: "Arial",
                  bold: true,
                  size: 20,
                }),
                new TextRun({
                  text: "\tValorización de Prestaciones ……..(Precisar según corresponda el tipo de prestación: ",
                  font: "Arial",
                  size: 19,
                }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "\t\tCondición de reembolso, sin cobertura, tercero o EPS) ",
                  font: "Arial",
                  size: 19,
                }),
              ],
              spacing: { after: 80 },
            }),

            // Referencia
            new Paragraph({
              children: [
                new TextRun({
                  text: "Referencia: ",
                  font: "Arial",
                  bold: true,
                  size: 20,
                }),
                new TextRun({
                  text: "\t......................... (N° de Informe de Auditoría de Seguros)",
                  font: "Arial",
                  size: 20,
                }),
              ],
              spacing: { after: 80 },

            }),

            // Exp
            new Paragraph({
              children: [
                new TextRun({
                  text: "Expediente: ",
                  font: "Arial",
                  bold: true,
                  size: 20,
                }),
              ],
              spacing: { after: 80 },

            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "Presente. -",
                  font: "Arial",
                  bold: true,
                  size: 20,
                  underline: { type: "single" },
                }),
              ],
              spacing: { after: 300 },
            }),

            // Texto principal
            new Paragraph({
              spacing: { after: 200 },
              children: [
                new TextRun({
                  text: "Atención: Unidad/Oficina/División de Finanzas",
                  font: "Arial",
                  bold: true,
                  size: 20,
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.JUSTIFIED,
              spacing: { after: 200 },
              children: [
                new TextRun({
                  text: "Tengo el agrado de dirigirme a Usted para saludarlo(a) cordialmente y de acuerdo al asunto de la referencia, solicitarle tenga a bien disponer la ",
                  font: "Arial",
                  size: 20,
                }),
                new TextRun({
                  text: "valorización de las prestaciones indebidamente otorgadas / en condición de reembolso,",
                  font: "Arial",
                  bold: true,
                  size: 20,
                }),
                new TextRun({
                  text: " de acuerdo a los documentos adjuntos.",
                  font: "Arial",
                  size: 20,
                }),
              ],
            }),

            // Normativa
            new Paragraph({
              alignment: AlignmentType.JUSTIFIED,
              spacing: { after: 200 },
              children: [
                new TextRun({
                  text: "En ese sentido, de acuerdo a lo establecido en la ",
                  font: "Arial",
                  size: 20,
                }),
                new TextRun({
                  text: "normatividad de la materia, ",
                  font: "Arial",
                  bold: true,
                  size: 20,
                }),
                new TextRun({
                  text: "agradeceremos nos informe los resultados de la valorización y acciones de cobranza dispuesto para las prestaciones otorgadas, en los plazos establecidos para esta actividad en la Directiva N° 14-GG-ESSALUD-2011 (prestaciones indebidas) o el TUO de la Ley N° 27444, aprobada mediante D.S. N° 004-2019-JUS (empleador con condición de reembolso de trabajadores del hogar).",
                  font: "Arial",
                  size: 20,
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.JUSTIFIED,
              spacing: { after: 200 },
              children: [
                new TextRun({
                  text: "Sin otro particular, agradezco anticipadamente su atención.",
                  font: "Arial",
                  size: 20,
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.JUSTIFIED,
              spacing: { after: 2000 },
              children: [
                new TextRun({
                  text: "Atentamente, ",
                  font: "Arial",
                  size: 20,
                }),
              ],
            }),

            // Firma
            new Paragraph({
              alignment: AlignmentType.JUSTIFIED,
              children: [
                new TextRun({
                  text: "\t______________________________",
                  font: "Arial",
                  size: 16,
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.JUSTIFIED,
              spacing: { after: 1600},
              children: [
                new TextRun({
                  text: "\t    Firma y Sello del Jefe de OSPE",
                  font: "Arial",
                  size: 16,
                }),
              ],
            }),

            // Footer
            new Paragraph({
              children: [
                new TextRun({
                  text: "XXX/xxx/xxx",
                  font: "Arial",
                  size: 18,
                  bold: true,
                }),
              ],
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "NIT N°",
                  font: "Arial",
                  size: 18,
                  bold: true,
                }),
              ],
              spacing: { after: 150},
            }),
            new Paragraph({
              children: [
                new TextRun({
                  text: "Cc: Finanzas",
                  font: "Arial",
                  size: 18,
                  bold: true,
                }),
              ],
            }),
            new Paragraph({
              spacing: { after: 15000 },
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 200 },
              children: [
                new TextRun({
                  text: "ANEXO",
                  font: "Arial",
                  size: 24,
                  bold: true,
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 200 },
              children: [
                new TextRun({
                  text: "Solicitud de Valorización de Prestaciones de Salud a Entidades Empleadoras en Condición de Reembolso",
                  font: "Arial",
                  size: 24, // 12 puntos
                  bold: true,
                }),
              ],
            }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                // Encabezado
                new TableRow({
                  children: [
                    new TableCell({
                      width: { size: 40, type: WidthType.PERCENTAGE },
                      borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.SINGLE } },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: "Características",
                              font: "Arial",
                              size: 22,
                              bold: true,
                            }),
                          ],
                          spacing: { after: 50 },
                        }),
                      ],
                    }),
                    new TableCell({
                      width: { size: 60, type: WidthType.PERCENTAGE },
                      borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.SINGLE } },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: "Detalle",
                              font: "Arial",
                              size: 22,
                              bold: true,
                            }),
                          ],
                          spacing: { after: 50 },
                        }),
                      ],
                    }),
                  ],
                }),
                ...[
                  ["Apellidos y nombres del asegurado", "«Nombres_del_auditado» «Apellido_Paterno» «Apellido_Materno»"],
                  ["Tipo de documento de Identidad", "«DNI»"],
                  ["N° De Documento de Identidad", "«XXXXXXXX»"],
                  ["Nombre del Titular y Documento de identidad (llenar sólo si el paciente es un Derechohabiente)", "«Nombres» «Apellido_Paterno» «Apellido_Materno»"],
                  ["Tipo de seguro", "«Tipo_de_Seguro»"],
                  ["Tipo de asegurado", "«Tipo_de_Asegurado»"],
                  ["IPRESS donde ocurrió la prestación", "«IPRESS»"],
                  ["Diagnostico(s)", "«CIE_10» «Diagnóstico_de_Evaluación»"],
                  ["Entidad empleadora auditada", "«Nombre_o_Razón_Social_del_Empleador»"],
                  ["N° RUC del empleador", "«RUC_del_Empleador»"],
                  ["Periodo de evaluación del indicador de condición de reembolso (18 meses)", "MM/AAAA – MM/AAAA"],
                  ["Mes(es) a valorizar", "Mes"],
                  ["Norma vulnerada", "«Norma_Vulnerada»"],
                ].map(([label, value]) =>
                  new TableRow({
                    children: [
                      new TableCell({
                        width: { size: 40, type: WidthType.PERCENTAGE },
                        borders: { bottom: { style: BorderStyle.SINGLE } },
                        children: [
                          new Paragraph({
                            indent: { left: 100 },
                            children: [
                              new TextRun({
                                text: label,
                                font: "Arial",
                                size: 20,
                              }),
                            ],
                            spacing: { after: 80 },
                          }),
                        ],
                      }),
                      new TableCell({
                        width: { size: 60, type: WidthType.PERCENTAGE },
                        borders: { bottom: { style: BorderStyle.SINGLE } },
                        children: [
                          new Paragraph({
                            indent: { left: 100 },
                            children: [
                              new TextRun({
                                text: value,
                                font: "Arial",
                                size: 20,
                              }),
                            ],
                            spacing: { after: 80 },
                          }),
                        ],
                      }),
                    ],
                  })
                ),
              ],
            }),
            new Paragraph({
              spacing: { after: 15000 },
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 200 },
              children: [
                new TextRun({
                  text: "ANEXO",
                  font: "Arial",
                  size: 24,
                  bold: true,
                }),
              ],
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              spacing: { after: 200 },
              children: [
                new TextRun({
                  text: "Solicitud de Valorización de Prestaciones de Salud No Coberturadas",
                  font: "Arial",
                  size: 24, // 12 puntos
                  bold: true,
                }),
              ],
            }),
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                // Encabezado de la tabla
                new TableRow({
                  children: [
                    new TableCell({
                      width: { size: 40, type: WidthType.PERCENTAGE },
                      borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.SINGLE } },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: "Características",
                              font: "Arial",
                              size: 22,
                              bold: true,
                            }),
                          ],
                          spacing: { after: 50 },
                        }),
                      ],
                    }),
                    new TableCell({
                      width: { size: 60, type: WidthType.PERCENTAGE },
                      borders: { top: { style: BorderStyle.SINGLE }, bottom: { style: BorderStyle.SINGLE } },
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: "Detalle",
                              font: "Arial",
                              size: 22,
                              bold: true,
                            }),
                          ],
                          spacing: { after: 50 },
                        }),
                      ],
                    }),
                  ],
                }),
                ...[
                  ["Apellidos y nombres del paciente", "«Nombres_del_auditado» «Apellido_Paterno» «Apellido_Materno»"],
                  ["Tipo de documento de Identidad", "«DNI»"],
                  ["N° De Documento de Identidad", "  «XXXXXXXX»"],
                  ["Nombre del Titular y Documento de   Identidad (llenar sólo si el paciente es   un Derechohabiente)", "«Nombres» «Apellido_Paterno» «Apellido_Materno»"],
                  ["Tipo de seguro", "«Tipo_de_Seguro» - «En_caso_de_otro_especificar1»"],
                  ["Tipo de asegurado", "«Tipo_de_Asegurado»"],
                  ["Tipo de Contingencia", "«Tipo_de_Contingencia»"],
                  ["Fecha de inicio y fin de la contingencia   (Periodo a valorizar)", "«FechaPeriodo_de_Contingencia»"],
                  ["IPRESS donde ocurrió la prestación", "«IPRESS»"],
                  ["Diagnóstico No Coberturado", "«CIE_10» «Diagnóstico_de_Evaluación»"],
                  ["Procedimiento no coberturado", "«Procedimiento no coberturado»"],
                  ["Norma vulnerada", "«Norma vulnerada»"],
                ].map(([label, value]) =>
                  new TableRow({
                    children: [
                      new TableCell({
                        width: { size: 40, type: WidthType.PERCENTAGE },
                        borders: { bottom: { style: BorderStyle.SINGLE } },
                        children: [
                          new Paragraph({
                            indent: { left: 100 },
                            alignment: AlignmentType.LEFT,
                            children: [
                              new TextRun({
                                text: label,
                                font: "Arial",
                                size: 20, // 10 puntos
                              }),
                            ],
                            spacing: { after: 80 },
                          }),
                        ],
                      }),
                      new TableCell({
                        width: { size: 60, type: WidthType.PERCENTAGE },
                        borders: { bottom: { style: BorderStyle.SINGLE } },
                        children: [
                          new Paragraph({
                            indent: { left: 100 },
                            alignment: AlignmentType.LEFT,
                            children: [
                              new TextRun({
                                text: value,
                                font: "Arial",
                                size: 20,
                              }),
                            ],
                            spacing: { after: 80 },
                          }),
                        ],
                      }),
                    ],
                  })
                ),
              ],
            }),
            new Paragraph({
              spacing: { after: 15000 },
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "ANEXO",
                  font: 'Arial',
                  size: 24,
                  bold: true,
                }),
              ],
              spacing: { after: 300 },
            }),
            new Paragraph({
              alignment: AlignmentType.CENTER,
              children: [
                new TextRun({
                  text: "Solicitud de Valorización de Prestaciones de Salud No Coberturadas a EPS",
                  font: 'Arial',
                  bold: true,
                  size: 24,
                }),
              ],
              spacing: { after: 200 },
            }),

            // Tabla
            new Table({
              width: { size: 100, type: WidthType.PERCENTAGE },
              rows: [
                // Fila de encabezado
                new TableRow({
                  children: [
                    new TableCell({
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: "Características",
                              font: "Arial",
                              size: 22,
                              bold: true,
                            }),
                          ],
                          spacing: { after: 50 },
                        }),
                      ],
                      borders: {
                        top: { style: BorderStyle.SINGLE },
                        bottom: { style: BorderStyle.SINGLE },
                      },
                    }),
                    new TableCell({
                      children: [
                        new Paragraph({
                          alignment: AlignmentType.CENTER,
                          children: [
                            new TextRun({
                              text: "Detalle",
                              font: "Arial",
                              size: 22,
                              bold: true,
                            }),
                          ],
                          spacing: { after: 50 },

                        }),
                      ],
                      borders: {
                        top: { style: BorderStyle.SINGLE },
                        bottom: { style: BorderStyle.SINGLE },
                      },
                    }),
                  ],
                }),
                ...[
                  ["Apellidos y nombres del asegurado", "«Nombres_del_auditado» «Apellido_Paterno» «Apellido_Materno»"],
                  ["Tipo de documento de Identidad", "«DNI»"],
                  ["N° De Documento de Identidad", "«XXXXXXXX»"],
                  ["Nombre del Titular y Documento de   identidad (llenar sólo si el paciente es   un Derechohabiente)", "«Nombres» «Apellido_Paterno» «Apellido_Materno»"],
                  ["Tipo de seguro", "«Tipo_de_Seguro» - «En_caso_de_otro_especificar1»"],
                  ["Tipo de asegurado", "«Tipo_de_Asegurado»"],
                  ["EPS afiliada", "«Nombre de la EPS»"],
                  ["RUC de la EPS afiliada", " «RUC de la EPS»"],
                  ["Tipo de Contingencia", "«Tipo_de_Contingencia»"],
                  ["Fecha de inicio y fin de la contingencia   (Periodo a valorizar)", "«FechaPeriodo_de_Contingencia»"],
                  ["IPRESS donde ocurrió la prestación.", "«IPRESS»"],
                  ["Diagnóstico No Coberturado", "«CIE_10» «Diagnóstico_de_Evaluación»"],
                  ["Procedimiento no coberturado", "«Procedimiento no coberturado»"],
                  ["Norma vulnerada", "«Norma vulnerada»"],
                ].map(([label, value]) =>
                  new TableRow({
                    children: [
                      new TableCell({
                        width: { size: 40, type: WidthType.PERCENTAGE },
                        children: [
                          new Paragraph({
                            indent: { left: 100 },
                            children: [
                              new TextRun({
                                text: label,
                                font: "Arial",
                                size: 20,
                              }),
                            ],
                            spacing: { after: 80 },
                          }),
                        ],
                        borders: { bottom: { style: BorderStyle.SINGLE } },
                      }),
                      new TableCell({
                        width: { size: 60, type: WidthType.PERCENTAGE },
                        children: [
                          new Paragraph({
                            indent: { left: 100 },
                            children: [
                              new TextRun({
                                text: value,
                                font: "Arial",
                                size: 20,
                              }),
                            ],
                            spacing: { after: 80 },
                          }),
                        ],
                        borders: { bottom: { style: BorderStyle.SINGLE } },
                      }),
                    ],
                  })
                ),
              ],
            }),
          ],
        },
      ],
    });
    Packer.toBlob(doc).then((blob) => {
      saveAs(blob, "informe_auditoria.docx");
    });
  }

}

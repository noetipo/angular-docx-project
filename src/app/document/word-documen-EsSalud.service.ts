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
  TabStopPosition, BorderStyle
} from 'docx';
import { saveAs } from 'file-saver';

@Injectable({
  providedIn: 'root'
})
export class DocumentEssaludService {
  generateDocument(): void {
    const defaultParagraphSpacing = {spacing: {before: 200, after: 200}};

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
                  // Espacio inicial en el encabezado
                  new Paragraph({
                    spacing: { after: 250 },
                  }),
                  // Texto centrado
                  new Paragraph({
                    alignment: AlignmentType.CENTER,
                    children: [
                      new TextRun({
                        text: "Año del Bicentenario, de la consolidación de nuestra Independencia, y de la conmemoración de las heroicas batallas de Junín y Ayacucho",
                        font: "Arial",
                        size: 16, // Tamaño de 8 puntos (16 half-points)
                      }),
                    ],
                    spacing: { after: 250 },
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
                    alignment: AlignmentType.LEFT,
                    children: [
                      new TextRun({
                        text: "\t\t           Jr. Domingo Cueto N.º 120   ",
                        font: "Arial",
                        size: 14,
                        bold: true,
                      }),
                    ],
                  }),
                  new Paragraph({
                    alignment: AlignmentType.LEFT,
                    children: [
                      new TextRun({
                        text: "www.gob.pe/essalud",
                        font: "Arial",
                        size: 14,
                        bold: true,

                      }),
                      new TextRun({
                        text: "           Jesús Maria",
                        font: "Arial",
                        size: 14,
                        bold: true,
                      }),
                    ],
                  }),

                  new Paragraph({
                    alignment: AlignmentType.LEFT,
                    children: [
                      new TextRun({
                        text: "\t\t           Lima 11 - Perú   ",
                        font: "Arial",
                        size: 14,
                        bold: true,
                      }),
                    ],
                  }),
                  new Paragraph({
                    alignment: AlignmentType.LEFT,
                    children: [
                      new TextRun({
                        text: "\t\t           Tel.:265 - 6000 / 265 - 7000   ",
                        font: "Arial",
                        size: 14,
                        bold: true,
                      }),
                    ],
                    spacing: { after: 200 },
                  }),
                ],
              }),
            },
            children: [
              new Paragraph({
                spacing: { after: 100 },
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "INFORME Nº …. -………..- …….-………..-ESSALUD-20XX",
                    bold: true,
                    font: "Arial",
                  }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: {after: 100},
              }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\t          Lima, …  de … de …",
                    font: "Arial",
                    bold: true,
                  }),
                ],
                alignment: AlignmentType.LEFT,
                spacing: {after: 400},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Para: \t\t«NOMBRE_DE_JEFE_OSPE»",
                    bold: true,
                    font: "Arial",
                  }),
                ],
                spacing: {after: 0},
              }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\t«JEFE DE LA OFICINA DE SEGUROS Y PRESTACIONES ECONÓMICAS «OSPE»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 320},
              }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "De: \t\t«NOMBRE_DE_AUDITOR»",
                    bold: true,
                    font: "Arial",
                  }),
                ],
                spacing: {after: 0},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\t«AUDITOR DE SEGUROS LA OFICINA DE SEGUROS Y PRESTACIONES                                       \t\tECONÓMICAS «NOMBRE DE OFICINA»",
                    font: "Arial",
                  }),
                ],
                alignment: AlignmentType.JUSTIFIED,
                spacing: {after: 320},
              }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "Asunto:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\tAuditoría de Seguros N°",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 120},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Expediente:",
                    bold: true,
                    font: "Arial",
                  }),
                ],
                spacing: {after: 120},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Referencia:",
                    bold: true,
                    font: "Arial",
                  }),
                ],
                spacing: {after: 200},
                border: {
                  bottom: {
                    color: "000000",
                    space: 108,
                    style: "single",
                    size: 6,
                  },
                },
              }),


              // Cuerpo del documento
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Mediante el presente me dirijo a usted para informar lo siguiente:",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 300},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "    ANTECEDENTES",
                    bold: true,
                    font: "Arial",
                  }),
                ],
                spacing: {after: 200},
                numbering: {reference: "numbering2", level: 0}, // Nivel 0
              }),
              new Paragraph({
                text: "",
                numbering: {reference: "numbering1", level: 1}, // Nivel 1
              }),
              new Paragraph({
                text: "",
                numbering: {reference: "numbering1", level: 1},
              }),
              new Paragraph({
                text: "",
                numbering: {reference: "numbering1", level: 1},
              }),
              new Paragraph({
                text: "",
                numbering: {reference: "numbering1", level: 1},
                spacing: {after: 300},

              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "    ANÁLISIS",
                    bold: true,
                    font: "Arial",
                  }),
                ],
                spacing: {after: 200},
                numbering: {reference: "numbering2", level: 0}, // Nivel 0
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "DATOS DE LA AUDITORÍA:",
                    bold: true,
                    font: "Arial",
                  }),
                ],
                numbering: {reference: "numbering3", level: 1},
                spacing: {after: 150},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tTipos de Auditoria:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t\t«Tipo_de_Auditoría»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 50},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tIPRESS:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t\t\t«IPRESS»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 50},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tFecha de Auditoría:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t\t«Fecha_de_Auditoría_de_Seguros»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 50},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tFecha de Contingencia:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t«FechaPeriodo_de_Contingencia»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 50},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tServicio:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t\t\t«Servicio_Asistencial_de_laContingencia»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 150},
              }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "DATOS DEL AUDITADO:",
                    bold: true,
                    font: "Arial",
                  }),
                ],
                numbering: {reference: "numbering3", level: 1},
                spacing: {after: 150},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tApellidos y Nombres:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t\t«Apellido_Paterno» «Apellido_Materno», «Nombres»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 50},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tTipo de Documento:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t\t«Tipo_de_Documento_de_Identidad»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 50},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tN° de Documento:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t\t«N_de_Documento_de_Identidad»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 50},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tCondición del Asegurado:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t«Condición_del_Asegurado»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 50},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tSituación del Asegurado:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t«Situación_del_Asegurado»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 50},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tAcreditado:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t\t\t«Acreditado»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 50},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tTipo de Seguro:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t\t«Tipo_de_Seguro»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 150},
              }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "DATOS DEL EMPLEADOR:",
                    bold: true,
                    font: "Arial",
                  }),
                ],
                numbering: {reference: "numbering3", level: 1},
                spacing: {after: 150},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tNombre o Razón Social:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t«Nombre_o_Razón_Social_del_Empleador»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 50},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tRUC:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t\t\t\t«RUC_del_Empleador»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 50},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tCondición del Empleador:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t«Condición_del_Empleador»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 300},
              }),
              new Paragraph({
                spacing: { after: 100 },
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "DATOS DE LA CONTINGENCIA:",
                    bold: true,
                    font: "Arial",
                  }),
                ],
                numbering: {reference: "numbering3", level: 1},
                spacing: { after: 150 },
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tTipo de Contingencia:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t\t«Tipo_de_Contingencia»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 50},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tDiagnóstico:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t\t\t«Diagnóstico_COBERTURADO/NO COBERTURADO»",
                    font: "Arial",
                    size: 18,
                  }),
                ],
                spacing: {after: 300},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "\t\tProcedimiento:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t\t«Diagnóstico_COBERTURADO/NO ",
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t\t\t\tCOBERTURADO»",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 300},
              }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "NO CONFORMIDADES:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t «No_Conformidades»",
                    font: "Arial",
                  }),
                ],
                numbering: {reference: "numbering3", level: 1},
                spacing: {after: 150},
              }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "OBSERVACIONES:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "\t\t «Observaciones»",
                    font: "Arial",
                  }),
                ],
                numbering: {reference: "numbering3", level: 1},
                spacing: {after: 150},
              }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "NORMATIVIDAD VULNERADA:",
                    bold: true,
                    font: "Arial",
                  }),
                  new TextRun({
                    text: "«Norma_Vulnerada»",
                    font: "Arial",
                  }),
                ],
                numbering: {reference: "numbering3", level: 1},
                spacing: {after: 150},
              }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "    CONCLUSIONES:",
                    bold: true,
                    font: "Arial",
                  }),
                ],
                spacing: {after: 200},
                numbering: {reference: "numbering2", level: 0},
              }),
              new Paragraph({
                text: "",
                numbering: {reference: "numbering4", level: 1},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "",
                    font: "Arial",
                  }),
                ],
                numbering: {reference: "numbering4", level: 1},
                spacing: {after: 400},

              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "    RECOMENDACIONES:",
                    bold: true,
                    font: "Arial",
                  }),
                ],
                spacing: {after: 200},
                numbering: {reference: "numbering2", level: 0},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "«Recomendaciones»",
                    font: "Arial",
                  }),
                ],
                numbering: {reference: "numbering5", level: 1},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "",
                  }),
                ],
                numbering: {reference: "numbering5", level: 1},
                spacing: {after: 150},
              }),
              new Paragraph({
                children: [
                  new TextRun({
                    text: "Es todo por cuanto informo para los fines que estime pertinente.",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 200},
              }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "Atentamente",
                    font: "Arial",
                  }),
                ],
                spacing: {after: 900},
              }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "__________________________________________",
                    font: "Arial",
                    bold: true,
                  }),
                ],
                alignment: AlignmentType.CENTER,
                spacing: {after: 100},
              }),

              new Paragraph({
                children: [
                  new TextRun({
                    text: "Firma y sello auditor de seguros de OSPE",
                    font: "Arial",
                  }),
                ],
                alignment: AlignmentType.CENTER,
              }),
            ],
          },
        ],
      });

      // Generar y descargar el documento
      Packer.toBlob(doc).then((blob) => {
        saveAs(blob, "informe_auditoria.docx");
      });
    }

}

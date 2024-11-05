// src/app/services/document.service.ts

import { Injectable } from '@angular/core';
import { Document, Packer, Paragraph, Table, TableRow, TableCell, TextRun, WidthType, AlignmentType, HeadingLevel, BorderStyle, ShadingType, TableLayoutType } from 'docx';
import { saveAs } from 'file-saver';

@Injectable({
  providedIn: 'root'
})
export class DocumentService {

  generateDocument(): void {
    const doc = new Document({
      sections: [
        {
          children: [
            // Título del Documento
            new Paragraph({
              text: 'ANEXO N° 02: “Orden de Verificación”',
              alignment: AlignmentType.CENTER,
              heading: HeadingLevel.HEADING_1,
            }),

            // Tabla de encabezado con párrafo de introducción y tabla anidada
            this.createHeaderTableWithIntro(),

            // Texto después de la tabla en un cuadro
            this.createTextInBox(),
            this.createDocumentsTableWithEmptyMiddleColumn(),
            this.createVerificationInfoBox(),
            // Firma del jefe
            new Paragraph({
              text: "Firma y Sello del Jefe Unidad de Control de las Filtraciones u OSPE, según el tipo de oficina",
              alignment: AlignmentType.CENTER,
              spacing: { before: 300 },
            })
          ]
        }
      ]
    });

    // Genera y descarga el documento
    Packer.toBlob(doc).then((blob) => {
      saveAs(blob, 'ANEXO_02_Orden_de_Verificacion.docx');
    });
  }

  private createHeaderTableWithIntro(): Table {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph("Orden de Verificación N°")],
              columnSpan: 2,
              margins: { top: 100, bottom: 100, left: 100, right: 100 },
            }),
            new TableCell({
              children: [new Paragraph("Ciudad y Fecha:")],
              columnSpan: 2,
              margins: { top: 100, bottom: 100, left: 100, right: 100 },
            })
          ]
        }),
        new TableRow({
          children: [new TableCell({ children: [new Paragraph("Razón Social/ Apellidos y Nombres:")], columnSpan: 4, margins: { top: 100, bottom: 100, left: 100, right: 100 } })]
        }),
        new TableRow({
          children: [new TableCell({ children: [new Paragraph("Tipo y Número de Documento:")], columnSpan: 4, margins: { top: 100, bottom: 100, left: 100, right: 100 } })]
        }),
        new TableRow({
          children: [new TableCell({ children: [new Paragraph("Domicilio real () / fiscal ()")], columnSpan: 4, margins: { top: 100, bottom: 100, left: 100, right: 100 } })]
        }),
        // Nueva fila para el párrafo de introducción y tabla anidada
        new TableRow({
          children: [
            new TableCell({
              children: [
                // Párrafo de introducción
                new Paragraph({
                  text: "Presente.-\nNos dirigimos a Ud., para comunicarle que, de acuerdo a lo establecido en la Ley N° 29135, modificada por Decreto Legislativo N° 1172, y su Reglamento aprobado por Decreto Supremo N° 002-2009-TR, se dará inicio al procedimiento de verificación del(los) siguiente(s) asegurado(s):",
                  spacing: { after: 300 },
                }),
                // Tabla anidada con diseño personalizado
                new Table({
                  width: { size: 100, type: WidthType.PERCENTAGE },
                  layout: TableLayoutType.FIXED,
                  rows: [
                    // Fila de encabezado
                    new TableRow({
                      children: [
                        new TableCell({
                          children: [new Paragraph({ text: "N°", alignment: AlignmentType.CENTER })],
                          borders: this.setBorders(),
                          rowSpan: 2,
                          margins: { top: 100, bottom: 100, left: 100, right: 100 },
                        }),
                        new TableCell({
                          children: [new Paragraph({ text: "Nombres y Apellidos", alignment: AlignmentType.CENTER })],
                          borders: this.setBorders(),
                          rowSpan: 2,
                          shading: { type: ShadingType.CLEAR, fill: "FFFF00" },
                          margins: { top: 100, bottom: 100, left: 100, right: 100 },
                        }),
                        new TableCell({
                          children: [new Paragraph({ text: "DNI", alignment: AlignmentType.CENTER })],
                          borders: this.setBorders(),
                          rowSpan: 2,
                          shading: { type: ShadingType.CLEAR, fill: "FFFF00" },
                          margins: { top: 100, bottom: 100, left: 100, right: 100 },
                        }),
                        new TableCell({
                          children: [new Paragraph({ text: "PERÍODO A VERIFICAR", alignment: AlignmentType.CENTER })],
                          borders: this.setBorders(),
                          columnSpan: 2,
                          margins: { top: 100, bottom: 100, left: 100, right: 100 },
                        })
                      ],
                    }),
                    // Subencabezados dentro de "PERÍODO A VERIFICAR"
                    new TableRow({
                      children: [
                        new TableCell({
                          children: [new Paragraph({ text: "Fecha de Inicio de la afiliación y/o inicio de relación laboral", alignment: AlignmentType.CENTER })],
                          borders: this.setBorders(),
                          margins: { top: 100, bottom: 100, left: 100, right: 100 },
                        }),
                        new TableCell({
                          children: [new Paragraph({ text: "Fecha fin de acreditación del asegurado", alignment: AlignmentType.CENTER })],
                          borders: this.setBorders(),
                          margins: { top: 100, bottom: 100, left: 100, right: 100 },
                        })
                      ]
                    }),
                    // Fila de datos de fechas
                    new TableRow({
                      children: [
                        new TableCell({
                          children: [new Paragraph("")],
                          borders: this.setBorders(),
                          shading: { type: ShadingType.CLEAR, fill: "FFFF00" },
                          margins: { top: 100, bottom: 100, left: 100, right: 100 },
                        }),
                        new TableCell({
                          children: [new Paragraph("")],
                          borders: this.setBorders(),
                          shading: { type: ShadingType.CLEAR, fill: "FFFF00" },
                          margins: { top: 100, bottom: 100, left: 100, right: 100 },
                        }),
                        new TableCell({
                          children: [new Paragraph("")],
                          borders: this.setBorders(),
                          shading: { type: ShadingType.CLEAR, fill: "FFFF00" },
                          margins: { top: 100, bottom: 100, left: 100, right: 100 },
                        }),
                        new TableCell({
                          children: [new Paragraph({ text: "dd/mm/aaaa", alignment: AlignmentType.CENTER })],
                          borders: this.setBorders(),
                          margins: { top: 100, bottom: 100, left: 100, right: 100 },
                        }),
                        new TableCell({
                          children: [new Paragraph({ text: "dd/mm/aaaa", alignment: AlignmentType.CENTER })],
                          borders: this.setBorders(),
                          margins: { top: 100, bottom: 100, left: 100, right: 100 },
                        })
                      ]
                    }),
                  ],
                })
              ],
              columnSpan: 4,
            })
          ]
        }),
      ],
    });
  }

  // Método para crear el cuadro alrededor del texto adicional
  private createTextInBox(): Table {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  text: "En virtud de lo dispuesto en el Artículo 11° del Reglamento aprobado por Decreto Supremo N° 002-2009-TR, le solicitamos poner a disposición del personal verificador los siguientes documentos:",
                  alignment: AlignmentType.LEFT,
                  spacing: { after: 300 }, // Espaciado similar al párrafo de introducción
                })
              ],
              borders: this.setBorders(),
              columnSpan: 4, // Combina ambas columnas para alineación consistente
              margins: { top: 100, bottom: 100, left: 100, right: 100 }, // Margen interno para el cuadro
            })
          ]
        })
      ],
    });
  }
  private createDocumentsTableWithEmptyMiddleColumn(): Table {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph("1) Contrato de Trabajo del asegurado")],
              borders: this.setBorders(),
              margins: { top: 100, bottom: 100, left: 100, right: 100 },
            }),
            new TableCell({ // Columna vacía combinada
              children: [new Paragraph("")],
              borders: this.setBorders(),
              rowSpan: 17, // Span para cubrir todas las filas en la columna del medio
              margins: { top: 100, bottom: 100, left: -1000, right: 2000 },
            }),
            new TableCell({
              children: [new Paragraph("10) Documentos de asignación de funciones del asegurado")],
              borders: this.setBorders(),
              margins: { top: 100, bottom: 100, left: 100, right: 100 },
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph("2) Declaraciones Tributarias-PDT- últimos seis (06) meses")],
              borders: this.setBorders(),
              margins: { top: 100, bottom: 100, left: 100, right: 100 },
            }),
            // Columna del medio vacía omitida ya que está combinada en la primera fila
            new TableCell({
              children: [new Paragraph("11) Documentos presentados por el trabajador al empleador que evidencien las funciones desarrolladas en los últimos seis (06) meses")],
              borders: this.setBorders(),
              margins: { top: 100, bottom: 100, left: 100, right: 100 },
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph("3) Registro en el MTPE")],
              borders: this.setBorders(),
              margins: { top: 100, bottom: 100, left: 100, right: 100 },
            }),
            new TableCell({
              children: [new Paragraph("12) Registro de Ventas. Comprobantes de pago que emite")],
              borders: this.setBorders(),
              margins: { top: 100, bottom: 100, left: 100, right: 100 },
            }),
          ]
        }),
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph("4) Registros especiales según su actividad económica")],
              borders: this.setBorders(),
              margins: { top: 100, bottom: 100, left: 100, right: 100 },
            }),
            new TableCell({
              children: [new Paragraph("13) Registro de Compras")],
              borders: this.setBorders(),
              margins: { top: 100, bottom: 100, left: 100, right: 100 },
            }),
          ]
        }),
        // Continúa de la misma manera para las filas restantes
      ],
    });
  }

  private createVerificationInfoBox(): Table {
    return new Table({
      width: { size: 100, type: WidthType.PERCENTAGE },
      rows: [
        new TableRow({
          children: [
            new TableCell({
              children: [
                new Paragraph({
                  text: "El verificador, de acuerdo a lo dispuesto en el Artículo 07° del Reglamento aprobado por Decreto Supremo N° 002-2009-TR, está facultado para iniciar la verificación inmediatamente después de recibida la Orden de Verificación, ingresar al centro de trabajo, levantar actas, practicar cualquier diligencia de investigación, examen o prueba que considere necesario, requerir información e identificación de las personas que se encuentren en el centro de trabajo materia de la acción de verificación y solicitar la comparecencia de la entidad empleadora o sus representantes, de los trabajadores y de cualesquiera sujetos incluidos en su ámbito de actuación en el centro inspeccionado.",
                  alignment: AlignmentType.JUSTIFIED,
                  spacing: { after: 300 },
                }),
                new Paragraph({
                  text: "El empleador debe permitir el ingreso a los funcionarios y/o servidores públicos en el centro de trabajo, lugar o establecimiento donde se lleva a cabo la verificación, colaborar con ellos durante su visita y facilitar la información y documentación que le sea solicitada para desarrollar la función de verificación.",
                  alignment: AlignmentType.JUSTIFIED,
                  spacing: { after: 300 },
                }),
                new Paragraph({
                  text: "Cabe precisar que el incumplimiento de lo señalado en el párrafo anterior constituye infracción tipificada en el Artículo 25° del Reglamento aprobado por Decreto Supremo N° 002-2009-TR, estando sujetos a las sanciones contenidas en el anexo de Tabla de Infracciones y Sanciones contenidas en el referido Decreto Supremo.",
                  alignment: AlignmentType.JUSTIFIED,
                  spacing: { after: 300 },
                }),
                new Paragraph({
                  text: "Si usted desea confirmar la identidad de los servidores designados podrá acceder a la siguiente dirección electrónica http://www.essalud.gob.pe/agencias-y-oficinas-de-seguros/ y/o comunicarse telefónicamente al número (Teléfono) y anexo de la Unidad de Control de las Filtraciones u OSPE, según el tipo de oficina) en el (horario de atención) para comprobar su identidad.",
                  alignment: AlignmentType.JUSTIFIED,
                  spacing: { after: 300 },
                }),
                new Paragraph({
                  text: "Base Legal: Ley N° 29135 reglamentada por el Decreto Supremo N° 002-2009-TR.",
                  alignment: AlignmentType.LEFT,
                })
              ],
              borders: this.setBorders(),
              columnSpan: 3, // Combina las tres columnas en una sola celda
              margins: { top: 100, bottom: 100, left: 100, right: 100 },
            })
          ]
        }),
        // Fila vacía
        new TableRow({
          children: [
            new TableCell({
              children: [new Paragraph("------------------------------------------------------------------------\n" +
                "Firma y Sello del Jefe Unidad de Control de las Filtraciones u OSPE, según el tipo de oficina\n")], // Celda vacía
              borders: this.setBorders(),
              columnSpan: 3, // Combina las tres columnas en una sola celda vacía
              margins: { top: 100, bottom: 100, left: 100, right: 100 },
            })
          ]
        }),
      ],
    });
  }





  // Método auxiliar para establecer bordes en las celdas
  private setBorders() {
    return {
      top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
    };
  }
}

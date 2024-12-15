import { Injectable } from '@angular/core';
import {
  Document,
  Packer,
  Paragraph,
  Table,
  TableRow,
  TableCell,
  TextRun,
  WidthType,
  AlignmentType,
  HeadingLevel,
  BorderStyle,
  ShadingType,
  TableLayoutType,
  VerticalAlign,
  Header,
  ImageRun,
} from 'docx';
import { saveAs } from 'file-saver';
import { logoEssalud } from './logo';

@Injectable({
  providedIn: 'root',
})
export class DocumentVerificationOrderService {
  generateDocument(data: any): void {
    // Configuración del estilo base para párrafos
    const defaultParagraphStyle = {
      spacing: { before: 120, after: 120, line: 360 },
      style: { font: { size: 24 } }, // Aumentar tamaño de letra (12pt = 24 half-points)
    };

    // Configuración de márgenes para celdas
    const defaultCellMargins = {
      top: 100,
      bottom: 100,
      left: 120,
      right: 120,
    };

    const base64Data = logoEssalud.split(',')[1];
    // Eliminar el prefijo 'data:image/png;base64,'
    const imageBytes = Uint8Array.from(atob(base64Data), (c) => c.charCodeAt(0));
    const header = new Header({
      children: [
        new Paragraph({
          children: [new ImageRun({ data: imageBytes, transformation: { width: 100, height: 32 } })],
          alignment: AlignmentType.LEFT,
        }),
      ],
    });

    const doc = new Document({
      sections: [
        {
          properties: {},
          headers: { default: header },
          children: [
            new Paragraph({
              children: [
                new TextRun({
                  text: 'ANEXO N° 02',
                  bold: true,
                }),
              ],
              alignment: AlignmentType.CENTER,
              spacing: { before: 100, after: 100 },
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
                          text: 'Orden de Verificación',
                          alignment: AlignmentType.CENTER,
                          spacing: { before: 100, after: 100 }, // Espaciado interno
                        }),
                      ],
                    }),
                  ],
                }),
              ],
            }),

            // Párrafo vacío para espacio entre tablas
            new Paragraph({
              text: '',
              spacing: { before: 100, after: 100 }, // Espacio entre tablas
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
                // Información general
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
                                        new TextRun({ text: 'Orden de Verificación:' }),
                                        new TextRun({
                                          text: ` N° ${data.ordenVerificacionNro}`,
                                          bold: true,
                                        }),
                                      ],
                                    }),
                                  ],
                                  width: { size: 50, type: WidthType.PERCENTAGE },
                                }),
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [new Paragraph('Ciudad y Fecha:')],
                                  width: { size: 50, type: WidthType.PERCENTAGE },
                                }),
                              ],
                            }),
                            new TableRow({
                              children: [
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [
                                    new Paragraph(
                                      `Razón Social/ Apellidos y Nombres: ${data.empresaRazonsocialApellidosNombres}`
                                    ),
                                  ],
                                  columnSpan: 2,
                                }),
                              ],
                            }),
                            new TableRow({
                              children: [
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [new Paragraph(`Tipo y Número de Documento: RUC ${data.empresaRuc}`)],
                                  columnSpan: 2,
                                }),
                              ],
                            }),
                            new TableRow({
                              children: [
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [
                                    new Paragraph(`Domicilio real () / fiscal (): ${data.empresaDomicilioRealFiscal}`),
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
                          text: 'Presente.-',
                          spacing: { before: 200, after: 200 },
                        }),
                        new Paragraph({
                          text: 'Nos dirigimos a Ud., para comunicarle que, de acuerdo a lo establecido en la Ley N° 29135, modificada por Decreto Legislativo N° 1172, y su Reglamento aprobado por Decreto Supremo N° 002-2009-TR, se dará inicio al procedimiento de verificación del(los) siguiente(s) asegurado(s):',
                          spacing: { after: 200 },
                        }),
                      ],
                    }),
                  ],
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
                                  children: [new Paragraph('N°')],
                                  width: { size: 10, type: WidthType.PERCENTAGE },
                                }),
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [new Paragraph('Nombres y Apellidos')],
                                  width: { size: 45, type: WidthType.PERCENTAGE },
                                }),
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [new Paragraph('DNI')],
                                  width: { size: 15, type: WidthType.PERCENTAGE },
                                }),
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [
                                    new Paragraph(
                                      'Período a Verificar – Fecha de Inicio de la afiliación hasta la fecha de acreditación del asegurado'
                                    ),
                                  ],
                                  width: { size: 30, type: WidthType.PERCENTAGE },
                                }),
                              ],
                            }),
                            new TableRow({
                              children: [
                                new TableCell({
                                  children: [new Paragraph('1')],
                                  shading: { type: ShadingType.CLEAR, fill: 'ffffff' },
                                  margins: { top: 100, bottom: 100, left: 100, right: 100 },
                                }),
                                new TableCell({
                                  children: [new Paragraph(`${data.aseguradoNombresApellidos}`)],
                                  shading: { type: ShadingType.CLEAR, fill: 'ffffff' },
                                  margins: { top: 100, bottom: 100, left: 100, right: 100 },
                                }),
                                new TableCell({
                                  children: [new Paragraph(`${data.aseguradoNumeroDocumento}`)],
                                  shading: { type: ShadingType.CLEAR, fill: 'ffffff' },
                                  margins: { top: 100, bottom: 100, left: 100, right: 100 },
                                }),
                                new TableCell({
                                  children: [
                                    new Paragraph({ text: 'Desde el aa/mm al aa/mm', alignment: AlignmentType.CENTER }),
                                  ],
                                  margins: { top: 100, bottom: 100, left: 100, right: 100 },
                                }),
                              ],
                            }),
                            ...Array(2)
                              .fill(null)
                              .map(
                                () =>
                                  new TableRow({
                                    children: [
                                      new TableCell({
                                        children: [new Paragraph('')],
                                        margins: { top: 100, bottom: 100, left: 100, right: 100 },
                                      }),
                                      new TableCell({
                                        children: [new Paragraph('')],
                                        margins: { top: 100, bottom: 100, left: 100, right: 100 },
                                      }),
                                      new TableCell({
                                        children: [new Paragraph('')],
                                        margins: { top: 100, bottom: 100, left: 100, right: 100 },
                                      }),
                                      new TableCell({
                                        children: [new Paragraph('Desde el aa/mm al aa/mm')],
                                        margins: { top: 100, bottom: 100, left: 100, right: 100 },
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
                // Texto verificadores y tabla
                new TableRow({
                  children: [
                    new TableCell({
                      children: [
                        new Paragraph({
                          text: 'Para llevar a cabo este procedimiento se ha designado a los siguientes verificadores:',
                          spacing: { before: 200, after: 200 },
                        }),
                        new Table({
                          width: { size: 100, type: WidthType.PERCENTAGE },
                          rows: [
                            new TableRow({
                              children: [
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [new Paragraph('N°')],
                                  width: { size: 10, type: WidthType.PERCENTAGE },
                                }),
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [new Paragraph('Nombres y Apellidos')],
                                  width: { size: 60, type: WidthType.PERCENTAGE },
                                }),
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [new Paragraph('DNI')],
                                  width: { size: 30, type: WidthType.PERCENTAGE },
                                }),
                              ],
                            }),
                            new TableRow({
                              children: [
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [new Paragraph('1')],
                                  width: { size: 10, type: WidthType.PERCENTAGE },
                                }),
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [new Paragraph(`${data.verificadorNombresApellidos}`)],
                                  width: { size: 60, type: WidthType.PERCENTAGE },
                                }),
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [new Paragraph(`${data.verificadorDni}`)],
                                  width: { size: 30, type: WidthType.PERCENTAGE },
                                }),
                              ],
                            }),
                            ...Array(1)
                              .fill(null)
                              .map(
                                () =>
                                  new TableRow({
                                    children: [
                                      new TableCell({ children: [new Paragraph('')] }),
                                      new TableCell({ children: [new Paragraph('')] }),
                                      new TableCell({ children: [new Paragraph('')] }),
                                    ],
                                  })
                              ),
                          ],
                        }),
                      ],
                    }),
                  ],
                }),
                // Texto documentos y tabla final
                new TableRow({
                  children: [
                    new TableCell({
                      children: [
                        // new Paragraph({
                        //   text: 'En virtud de lo dispuesto en el Artículo 11° del Reglamento aprobado por Decreto Supremo N° 002-2009-TR, le solicitamos poner a disposición del personal verificador los siguientes documentos:',
                        //   spacing: { before: 200, after: 200 },
                        //   bold: true,
                        // }),
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: 'En virtud de lo dispuesto en el Artículo 11° del Reglamento aprobado por Decreto Supremo N° 002-2009-TR, le solicitamos poner a disposición del personal verificador los siguientes documentos:',
                              // spacing: { before: 200, after: 200 },
                              bold: true,
                            }),
                          ],
                          spacing: { before: 100, after: 100 },
                        }),

                        new Table({
                          width: { size: 100, type: WidthType.PERCENTAGE },
                          rows: [
                            new TableRow({
                              children: [
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [
                                    new Paragraph('1) Contrato de Trabajo del asegurado'),
                                    new Paragraph('2) Declaraciones Tributarias-PDT 621- últimos seis (6) meses'),
                                    new Paragraph('3) Registro en el MTPE'),
                                    new Paragraph('4) Registros especiales según su actividad económica'),
                                    new Paragraph(
                                      '5) Planilla de Sueldos o Remuneraciones/Planilla Electrónica-PDT 601'
                                    ),
                                    new Paragraph('6) Boletas de Pago del asegurado de los últimos seis (6) meses'),
                                    new Paragraph('7) Partes diarios de Asistencia de los últimos seis (6) meses'),
                                    new Paragraph('8) Seis (6) últimos pagos de Aporte ONP/AFP del asegurado'),
                                    new Paragraph('9) Descansos médicos de los últimos seis (6) meses'),
                                    new Paragraph('10) Documentos de asignación de funciones del asegurado'),
                                  ],
                                  width: { size: 48, type: WidthType.PERCENTAGE },
                                }),
                                new TableCell({
                                  width: { size: 4, type: WidthType.PERCENTAGE },
                                  margins: defaultCellMargins,
                                  children: [new Paragraph('')],
                                }),
                                new TableCell({
                                  margins: defaultCellMargins,
                                  children: [
                                    new Paragraph(
                                      '11) Documentos presentados por el trabajador al empleador que evidencien las funciones desarrolladas en los últimos seis (6) meses'
                                    ),
                                    new Paragraph('12) Registro de Ventas'),
                                    new Paragraph('13) Comprobantes de pago que emite'),
                                    new Paragraph('14) Registro de Compras'),
                                    new Paragraph('15) Relación de Productos o Servicios que vende'),
                                    new Paragraph(
                                      '16) Carta de empleador para depósito de CTS de los últimos semestres'
                                    ),
                                    new Paragraph(
                                      '17) Título de propiedad, Título de Posesión, Contrato de Arrendamiento o Cesión de Uso del terreno donde se realiza la Actividad Agraria'
                                    ),
                                    new Paragraph('18) Puede considerar otros documentos'),
                                    new Paragraph(
                                      '19)	Otros: . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . . .'
                                    ),
                                    new Paragraph('Colocar el número de ítems solicitados: . . . . . . . .'),
                                  ],
                                  width: { size: 48, type: WidthType.PERCENTAGE },
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
                      children: [
                        new Paragraph({
                          text: 'El verificador, de acuerdo a lo dispuesto en el artículo 07° del Reglamento aprobado por Decreto Supremo N° 002-2009-TR, está facultado para iniciar la verificación inmediatamente después de recibida la Orden de Verificación, ingresar al centro de trabajo, levantar actas, practicar cualquier diligencia de investigación, examen o prueba que considere necesario, requerir información e identificación de las personas que se encuentren en el centro de trabajo materia de la acción de verificación y solicitar la comparecencia de la entidad empleadora o sus representantes, de los trabajadores y de cualesquiera sujetos incluidos en su ámbito de actuación en el centro inspeccionado.',
                          alignment: AlignmentType.JUSTIFIED,
                          spacing: { after: 100 },
                        }),
                        new Paragraph({
                          text: 'El empleador debe permitir el ingreso a los funcionarios y/o servidores públicos en el centro de trabajo, lugar o establecimiento donde se lleva a cabo la verificación, colaborar con ellos durante su visita y facilitar la información y documentación que le sea solicitada para desarrollar la función de verificación, el incumplimiento de lo señalado en el párrafo anterior constituye infracción tipificada en el artículo 25° del Reglamento en mención, estando sujetos a las sanciones contenidas en el anexo de Tabla de Infracciones y Sanciones contenidas en el referido Decreto Supremo.',
                          alignment: AlignmentType.JUSTIFIED,
                          spacing: { after: 100 },
                        }),
                        new Paragraph({
                          text: 'En caso el verificador no sea atendido o exista demora en la atención, llevará a cabo una nueva visita dentro de los tres (3) días hábiles siguientes, para lo cual se deberá tener toda la información y/o documentación a su disposición, tal como se señala en el artículo 16° de la norma antes acotada.',
                          alignment: AlignmentType.JUSTIFIED,
                          spacing: { after: 100 },
                        }),
                        new Paragraph({
                          text: 'Si usted desea confirmar la identidad de los servidores designados podrá acceder a la siguiente dirección electrónica http://www.essalud.gob.pe/agencias-y-oficinas-de-seguros/  y/o comunicarse telefónicamente al número (Teléfono y anexo de la UCF u OSPE, según el tipo de oficina) en el (horario de atención) para comprobar su identidad.',
                          alignment: AlignmentType.JUSTIFIED,
                          spacing: { after: 100 },
                        }),
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: 'Base Legal: Ley N° 29135 reglamentada por el Decreto Supremo N° 002-2009-TR.',
                              bold: true,
                            }),
                          ],
                          alignment: AlignmentType.LEFT,
                          // spacing: { before: 50, after: 100 },
                        }),
                      ],
                      columnSpan: 3, // Combina las tres columnas en una sola celda
                      margins: { top: 100, bottom: 100, left: 100, right: 100 },
                    }),
                  ],
                }),

                // new TableRow({
                //   children: [
                //     new TableCell({
                //       margins: defaultCellMargins,
                //       children: [
                //         new Paragraph({
                //           text: 'El verificador, de acuerdo a lo dispuesto en el artículo 07° del Reglamento aprobado por Decreto Supremo N° 002-2009-TR, está facultado para iniciar la verificación inmediatamente después de recibida la Orden de Verificación, ingresar al centro de trabajo, levantar actas, practicar cualquier diligencia de investigación, examen o prueba que considere necesario, requerir información e identificación de las personas que se encuentren en el centro de trabajo materia de la acción de verificación y solicitar la comparecencia de la entidad empleadora o sus representantes, de los trabajadores y de cualesquiera sujetos incluidos en su ámbito de actuación en el centro inspeccionado. El empleador debe permitir el ingreso a los funcionarios y/o servidores públicos en el centro de trabajo, lugar o establecimiento donde se lleva a cabo la verificación, colaborar con ellos durante su visita y facilitar la información y documentación que le sea solicitada para desarrollar la función de verificación, el incumplimiento de lo señalado en el párrafo anterior constituye infracción tipificada en el artículo 25° del Reglamento en mención, estando sujetos a las sanciones contenidas en el anexo de Tabla de Infracciones y Sanciones contenidas en el referido Decreto Supremo. En caso el verificador no sea atendido o exista demora en la atención, llevará a cabo una nueva visita dentro de los tres (3) días hábiles siguientes, para lo cual se deberá tener toda la información y/o documentación a su disposición, tal como se señala en el artículo 16° de la norma antes acotada. Si usted desea confirmar la identidad de los servidores designados podrá acceder a la siguiente dirección electrónica http://www.essalud.gob.pe/agencias-y-oficinas-de-seguros/  y/o comunicarse telefónicamente al número (Teléfono y anexo de la UCF u OSPE, según el tipo de oficina) en el (horario de atención) para comprobar su identidad. Base Legal: Ley N° 29135 reglamentada por el Decreto Supremo N° 002-2009-TR.',
                //           spacing: { before: 200, after: 200 },
                //         }),
                //       ],
                //       width: { size: 50, type: WidthType.PERCENTAGE },
                //     }),
                //   ],
                // }),

                new TableRow({
                  children: [
                    new TableCell({
                      margins: defaultCellMargins,
                      children: [
                        // new Paragraph({
                        //   text: 'Acepto que este documento y las demás comunicaciones me sean remitidas vía correo electrónico   SI    ',
                        //   spacing: { before: 200, after: 100 },
                        // }),
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: 'Acepto que este documento y las demás comunicaciones me sean remitidas vía correo electrónico: SI:',
                            }),
                            new TextRun({
                              text: '☐',
                              bold: true,
                              size: 24,
                            }),
                          ],
                          alignment: AlignmentType.JUSTIFIED,
                          spacing: { after: 300 },
                        }),
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: 'Correo:',
                              bold: true,
                            }),
                          ],
                          spacing: { before: 100, after: 100 },
                        }),
                        // new Paragraph({
                        //   text: 'Acepto la toma de fotos y videos en relación a mi caso de verificación.   SI    ',
                        //   spacing: { before: 100, after: 100 },
                        // }),

                        new Paragraph({
                          children: [
                            new TextRun({
                              text: 'Acepto la toma de fotos y videos en relación a mi caso de verificación. SI:',
                            }),
                            new TextRun({
                              text: '☐',
                              bold: true,
                              size: 24,
                            }),
                          ],
                          alignment: AlignmentType.JUSTIFIED,
                          spacing: { after: 300 },
                        }),
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: '_____________________________',
                              underline: {},
                            }),
                          ],
                          spacing: { before: 200, after: 50 },
                        }),
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: 'Firma y Sello del Jefe de OSPE',
                              bold: true,
                            }),
                          ],
                          alignment: AlignmentType.LEFT,
                          spacing: { before: 50, after: 100 },
                        }),
                        new Paragraph({
                          children: [
                            new TextRun({
                              text: 'Recepción: ............',
                              bold: true,
                            }),
                          ],
                          alignment: AlignmentType.RIGHT,
                          spacing: { before: 50, after: 100 },
                        }),
                      ],
                      width: { size: 50, type: WidthType.PERCENTAGE },
                    }),
                  ],
                }),
              ],
            }),
          ],
        },
      ],
    });

    // Generar y descargar el documento
    Packer.toBlob(doc).then((blob) => {
      saveAs(blob, 'Orden de Verificacion.docx');
    });
  }

  private base64ToArrayBuffer(base64: string): ArrayBuffer {
    const binaryString = window.atob(base64.split(',')[1]);
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
      bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;
  }
}

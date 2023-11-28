const express = require('express');
const router = express.Router();

const excel = require('exceljs');
const path = require('path');
const fs = require('fs/promises');


/* GET home page. */
router.get('/excel', async (req, res, next) => {
  try {
    // Crear un nuevo libro de Excel
    const workbook = new excel.Workbook();

    const sheetName = 'JUEVES 23';
    const worksheet = workbook.addWorksheet(sheetName);
    worksheet.properties.defaultRowHeight = 20;
    worksheet.properties.defaultColWidth = 14;

    const data = {
      title: 'PROGRAMACION QUIRURGICA JUEVES 16 DE NOVIEMBRE 2023',
      salas: [
        {
          "codigoSala": "SALA 01",
          items: [
            { hora: '06:00', anestesia: '', sala: 'I', paciente: 'ROSALES SANTIVAÑEZ, MIDA EUSEBIA', edad: 32, estancia: 'AMB', cx: 'MASTECTOMIA SIMPLE', cirujano: 'DR. MIGUEL PINILLOS', anestesiologo: '', insumos: '', tiempo: '2 HORAS' },
            { hora: '13:30', anestesia: '', sala: 'I', paciente: 'OBRIEN DURAN , VERONICA FABIOLA', edad: 46, estancia: 'HOSP', cx: 'CENS+TURBINO', cirujano: 'DR. CABRERA', anestesiologo: '', insumos: 'MICRODEBRILADOR+RADIOFRECUENCIA', tiempo: '3 HORAS' },
          ]
        },
        {
          "codigoSala": "SALA 02",
          items: [
            { hora: '07:00', anestesia: '', sala: 'II', paciente: 'YOVERA DE LA TORRE , CINTIA YESENIA', edad: 30, estancia: 'HOSP', cx: 'MICRODISECTOMIA DE HERNIA DISCAL', cirujano: 'DR. CONCHA', anestesiologo: '', insumos: 'ARCO EN C+MICROSCOPIO', tiempo: '2 HORAS' },
            { hora: '11:00', anestesia: '', sala: 'II', paciente: 'SILES GONZALES-VIGIL GUILLERMO', edad: 52, estancia: 'AMB', cx: 'ARTROSCOPIA DE RODILLA', cirujano: 'DR. VEGA', anestesiologo: '', insumos: '', tiempo: '4 HORAS' },
          ]
        },
        {
          "codigoSala": "SALA 03",
          items: [
            { hora: '08:00', anestesia: '', sala: 'III', paciente: 'YOVERA DE LA TORRE , CINTIA YESENIA', edad: 30, estancia: 'HOSP', cx: 'MICRODISECTOMIA DE HERNIA DISCAL', cirujano: 'DR. CONCHA', anestesiologo: '', insumos: 'ARCO EN C+MICROSCOPIO', tiempo: '2 HORAS' },
            { hora: '11:00', anestesia: '', sala: 'II', paciente: 'SILES GONZALES-VIGIL GUILLERMO', edad: 52, estancia: 'AMB', cx: 'ARTROSCOPIA DE RODILLA', cirujano: 'DR. VEGA', anestesiologo: '', insumos: '', tiempo: '4 HORAS' },
          ]
        },
        {
          "codigoSala": "SALA 04",
          items: [
            { hora: '08:30', anestesia: '', sala: 'IV', paciente: 'HERRERA BEGAZO , PATRICIA', edad: 57, estancia: 'AMB', cx: 'RESECCION RADICAL DE TUMOR DE PARTES BLANDOS', cirujano: 'DR. MIGUEL PINILLOS', anestesiologo: '', insumos: '', tiempo: '1 HORA' },
          ]
        },
      ],
    }

    console.log(data)


    // headers
    worksheet.columns = [
      { header: 'HORA', key: 'hora', width: 10, headerRow: 1 },
      { header: 'ANESTESIA', key: 'anestesia', width: 10, headerRow: 1 },
      { header: 'SALA', key: 'sala', width: 10, headerRow: 1 },
      { header: 'PACIENTE', key: 'paciente', width: 10, headerRow: 1 },
      { header: 'EDAD', key: 'edad', width: 10, headerRow: 1 },
      { header: 'ESTANCIA', key: 'estancia', width: 10, headerRow: 1 },
      { header: 'CX', key: 'CX', width: 10, headerRow: 1 },
      { header: 'CIRUJANO', key: 'cirujano', width: 10, headerRow: 1 },
      { header: 'ANESTESIOLOGO', key: 'anestesiologo', width: 10, headerRow: 1 },
      { header: 'INSUMOS', key: 'insumos', width: 10, headerRow: 1 },
      { header: 'TIEMPO QUIRÚRGICO', key: 'tiempo', width: 10, headerRow: 1 },
    ];
    const headerRow = worksheet.getRow(1);

    // title
    worksheet.insertRow(1, data.title);
    worksheet.mergeCells('A1:K1');
    const titleCell = worksheet.getCell('A1');
    titleCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FFFF00' } };
    titleCell.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };
    titleCell.font = { bold: true, size: 14 };
    titleCell.alignment = { horizontal: 'center' };
    titleCell.value = data.title;

    let rowNumber = 3;
    for (let sala of data.salas) {
      console.log({sala})
      for (let item of sala.items) {
        const horaCell = worksheet.getCell(`A${rowNumber}`);
        horaCell.value = item.hora;

        const salaCell = worksheet.getCell(`C${rowNumber}`);
        salaCell.value = item.sala;

        const pacienteCell = worksheet.getCell(`D${rowNumber}`);
        pacienteCell.value = item.paciente;

        const edadCell = worksheet.getCell(`E${rowNumber}`);
        edadCell.value = item.edad;

        const estanciaCell = worksheet.getCell(`F${rowNumber}`);
        estanciaCell.value = item.estancia;

        const cxCell = worksheet.getCell(`G${rowNumber}`);
        cxCell.value = item.cx;

        const cirujanoCell = worksheet.getCell(`H${rowNumber}`);
        cirujanoCell.value = item.cirujano;

        const insumosCell = worksheet.getCell(`J${rowNumber}`);
        insumosCell.value = item.insumos;

        const tiempoCell = worksheet.getCell(`K${rowNumber}`);
        tiempoCell.value = item.tiempo;


        rowNumber += 1;
      }
    }

    // const headerRow = worksheet.getRow(1);
    // headerRow.values = columnHeaders;
    // const headerRange = worksheet.getCell('A2:K2');
    // headerRange.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F4B084' } };
    // headerRange.font = { bold: true };
    // headerRange.alignment = { horizontal: 'center', wrapText: true };
    // headerRange.border = { top: { style: 'thin' }, left: { style: 'thin' }, bottom: { style: 'thin' }, right: { style: 'thin' } };


    // Crear un nombre de archivo único
    const filename = `documento_excel_${Date.now()}.xlsx`;

    // Guardar el libro de Excel en el servidor
    const filePath = path.join(__dirname, '..', 'excel', filename);

    await workbook.xlsx.writeFile(filePath);

    // Enviar el archivo como respuesta
    res.download(filePath, async (err) => {
      // Eliminar el archivo después de descargarlo
      if (!err) {
        // Puedes comentar la siguiente línea si prefieres conservar el archivo
        await fs.unlink(filePath);
      }
    });

  } catch (error) {
    console.error('Error al generar el archivo Excel:', error);
    res.status(500).send('Error interno del servidor');
  }

});

module.exports = router;

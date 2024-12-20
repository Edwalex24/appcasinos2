const express = require('express');
const fileUpload = require('express-fileupload');
const xlsx = require('xlsx');
const app = express();

app.use(fileUpload());
app.use(express.static('public'));

app.get('/', (req, res) => {
    res.sendFile(__dirname + '/public/index.html');
});

app.post('/upload', (req, res) => {
    if (!req.files || Object.keys(req.files).length < 2) {
        return res.status(400).send('Please upload both files.');
    }

    const file1 = req.files.file1;
    const file2 = req.files.file2;

    const workbook1 = xlsx.read(file1.data, { type: 'buffer' });
    const workbook2 = xlsx.read(file2.data, { type: 'buffer' });

    const sheetName1 = workbook1.SheetNames[0];
    const sheet1 = workbook1.Sheets[sheetName1];
    const range1 = xlsx.utils.decode_range(sheet1['!ref']);
    range1.s.r = 2;
    sheet1['!ref'] = xlsx.utils.encode_range(range1);
    const data1 = xlsx.utils.sheet_to_json(sheet1);

    const sheetName2 = workbook2.SheetNames[0];
    const sheet2 = workbook2.Sheets[sheetName2];
    const data2 = xlsx.utils.sheet_to_json(sheet2);

    const newSheetName = 'Anexo_procesado';
    const newSheetData = data1.map(row => ({
        'Contrato': row['Contrato'],
        'Cod local': row['Cod local'],
        'Establecimiento': row['Establecimiento'],
        'NUC': row['NUC'],
        'Serial': row['Serial'],
        'Código marca': row['Código marca'],
        'Tipo apuesta': row['Tipo apuesta'],
        'Fecha reporte': row['Fecha reporte'],
        'Base liquidación diaria': row['Base liquidación diaria'],
        'Cod_establecimiento1': `${row['Cod local']} ${row['Establecimiento']}`,
        '12% Derechos de Explotación': row['Base liquidación diaria'] * 0.12,
        '1% Gastos Administrativos': row['Base liquidación diaria'] * 0.12 * 0.01,
        'Neto1': (row['Base liquidación diaria'] * 0.12) + (row['Base liquidación diaria'] * 0.12 * 0.01)
    }));
    const newSheet = xlsx.utils.json_to_sheet(newSheetData);

    for (let i = 2; i <= newSheetData.length + 1; i++) {
        if (newSheet[`H${i}`]) newSheet[`H${i}`].z = 'dd/mm/yyyy';
        if (newSheet[`I${i}`]) newSheet[`I${i}`].z = '"$"#,##0';
        if (newSheet[`K${i}`]) newSheet[`K${i}`].z = '"$"#,##0';
        if (newSheet[`L${i}`]) newSheet[`L${i}`].z = '"$"#,##0';
        if (newSheet[`M${i}`]) newSheet[`M${i}`].z = '"$"#,##0';
    }

    xlsx.utils.book_append_sheet(workbook1, newSheet, newSheetName);

    const contabilizadosData = {};
    newSheetData.forEach(row => {
        const key = row['Cod_establecimiento1'];
        if (!contabilizadosData[key]) {
            contabilizadosData[key] = 0;
        }
        contabilizadosData[key] += row['Neto1'];
    });

    const contabilizadosArray = Object.keys(contabilizadosData).map(key => ({
        'Cod_establecimiento1': key,
        'Neto1': contabilizadosData[key]
    }));

    const totalNeto1 = contabilizadosArray.reduce((sum, row) => sum + row['Neto1'], 0);
    contabilizadosArray.push({
        'Cod_establecimiento1': 'Total a Pagar',
        'Neto1': totalNeto1
    });

    const contabilizadosSheet = xlsx.utils.json_to_sheet(contabilizadosArray);

    // Aquí es donde se aplican las mejoras de formato
    function applyCellStyles(sheet, range, styles) {
        if (!range || !range.s || !range.e) {
            console.error('Rango no definido correctamente:', range);
            return;
        }
        for (let R = range.s.r; R <= range.e.r; ++R) {
            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellAddress = xlsx.utils.encode_cell({ r: R, c: C });
                if (!sheet[cellAddress]) continue;
                sheet[cellAddress].s = styles;
            }
        }
    }

    const headerStyle = {
        font: { bold: true },
        alignment: { horizontal: 'center', vertical: 'center' },
        fill: { fgColor: { rgb: 'FFFF00' } },
        border: {
            top: { style: 'thin', color: { rgb: '000000' } },
            bottom: { style: 'thin', color: { rgb: '000000' } },
            left: { style: 'thin', color: { rgb: '000000' } },
            right: { style: 'thin', color: { rgb: '000000' } }
        }
    };

    const currencyStyle = {
        numFmt: '"$"#,##0',
        alignment: { horizontal: 'right' }
    };

    const headerRange = xlsx.utils.decode_range(contabilizadosSheet['!ref']);
    headerRange.e.r = 0;
    applyCellStyles(contabilizadosSheet, headerRange, headerStyle);

    const currencyRange = xlsx.utils.decode_range(contabilizadosSheet['!ref']);
    currencyRange.s.c = 1;
    currencyRange.e.c = 1;
    applyCellStyles(contabilizadosSheet, currencyRange, currencyStyle);

    const wscols = [
        { wch: 20 },
        { wch: 15 },
        { wch: 30 },
        { wch: 15 }
    ];
    contabilizadosSheet['!cols'] = wscols;

    contabilizadosSheet['!freeze'] = { xSplit: 1, ySplit: 1, topLeftCell: 'B2', activePane: 'bottomRight', state: 'frozen' };

    if (workbook1.SheetNames.includes('contabilizados')) {
        delete workbook1.Sheets['contabilizados'];
        workbook1.SheetNames = workbook1.SheetNames.filter(name => name !== 'contabilizados');
    }

    xlsx.utils.book_append_sheet(workbook1, contabilizadosSheet, 'contabilizados');

    const inventarioData = data2.map(row => ({
        ...row,
        'Nombre_Local_Inv': `${row['Código Local']} ${row['Nombre Establecimiento']}`
    }));

    const inventarioOcurrencias = inventarioData.reduce((acc, row) => {
        if (!acc[row['Nombre_Local_Inv']]) {
            acc[row['Nombre_Local_Inv']] = 0;
        }
        acc[row['Nombre_Local_Inv']] += 1;
        return acc;
    }, {});

    contabilizadosArray.forEach(row => {
        if (row['Cod_establecimiento1'] !== 'Total a Pagar') {
            row['Nombre_Local_Inv'] = row['Cod_establecimiento1'];
            row['Total Locales'] = inventarioOcurrencias[row['Cod_establecimiento1']] || 0;
        }
    });

    const totalMaquinas = contabilizadosArray.reduce((sum, row) => sum + (row['Total Locales'] || 0), 0);
    contabilizadosArray.push({
        'Cod_establecimiento1': '',
        'Neto1': '',
        'Nombre_Local_Inv': 'Cantidad de Maquinas',
        'Total Locales': totalMaquinas
    });

    const finalContabilizadosSheet = xlsx.utils.json_to_sheet(contabilizadosArray);

    if (workbook1.SheetNames.includes('contabilizados')) {
        delete workbook1.Sheets['contabilizados'];
        workbook1.SheetNames = workbook1.SheetNames.filter(name => name !== 'contabilizados');
    }

    xlsx.utils.book_append_sheet(workbook1, finalContabilizadosSheet, 'contabilizados');

    delete workbook1.Sheets[sheetName1];
    workbook1.SheetNames = workbook1.SheetNames.filter(name => name !== sheetName1);

    const buffer = xlsx.write(workbook1, { type: 'buffer', bookType: 'xlsx' });

    res.setHeader('Content-Disposition', 'attachment; filename=Anexo_procesado.xlsx');
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    res.send(buffer);
});

app.listen(3000, () => {
    console.log('Server started on http://localhost:3000');
});

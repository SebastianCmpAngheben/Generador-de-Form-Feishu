import React, { useEffect, useState } from 'react';
import { bitable } from '@lark-base-open/js-sdk';
import ExcelJS from 'exceljs';

interface GastoRecord {
  userCode: string;
  completeName: string;
  category: string;
  estimateBudget: number;
  amountReleased: number;
  uuidInvoice: string;
  fechaInicio: number | null;
  fechaTermino: number | null;
}

export default function App() {
  const [users, setUsers] = useState<string[]>([]);
  const [selectedUser, setSelectedUser] = useState<string>('');
  const [fechaInicioFiltro, setFechaInicioFiltro] = useState<string>('');
  const [fechaTerminoFiltro, setFechaTerminoFiltro] = useState<string>('');
  const [loading, setLoading] = useState<boolean>(false);
  const [reportData, setReportData] = useState<{ 
    userCode: string; 
    fullName: string; 
    records: GastoRecord[];
    totalBudget: number;
    totalReleased: number;
    minDate: number | null;
    maxDate: number | null;
    totalDays: number;
  } | null>(null);

  // --- PARSEADORES SEGUROS DE DATOS ---
  const parseTextField = (cellValue: any): string => {
    if (!cellValue) return '';
    if (typeof cellValue === 'string') return cellValue.trim();
    if (Array.isArray(cellValue)) {
      return cellValue.map((seg: any) => typeof seg === 'string' ? seg : (seg?.text || seg?.value || '')).join('').trim();
    }
    if (typeof cellValue === 'object') {
      if (cellValue.text !== undefined) return String(cellValue.text).trim();
      if (cellValue.value !== undefined) {
        if (typeof cellValue.value === 'string') return cellValue.value.trim();
        if (Array.isArray(cellValue.value)) return cellValue.value.join('').trim();
        return String(cellValue.value).trim();
      }
    }
    return cellValue.text ? String(cellValue.text).trim() : String(cellValue).trim();
  };

  const parseNumberField = (cellValue: any): number => {
    if (!cellValue) return 0;
    if (typeof cellValue === 'number') return cellValue;
    if (typeof cellValue === 'string') return parseFloat(cellValue.replace(/[^0-9.-]/g, '')) || 0;
    if (typeof cellValue === 'object') {
      const val = cellValue.value !== undefined ? cellValue.value : (Array.isArray(cellValue) ? cellValue[0] : cellValue.text);
      if (typeof val === 'number') return val;
      if (typeof val === 'string') return parseFloat(val.replace(/[^0-9.-]/g, '')) || 0;
      if (val && typeof val === 'object') return Number(val.text || val.value || 0);
    }
    return 0;
  };

  const parseDateField = (cellValue: any): number | null => {
    if (!cellValue) return null;
    if (typeof cellValue === 'number') return cellValue;
    if (Array.isArray(cellValue) && typeof cellValue[0] === 'number') return cellValue[0];
    if (typeof cellValue === 'object' && cellValue.value !== undefined) return Number(cellValue.value);
    return null;
  };

  const getFieldIdByNameSafe = async (table: any, fieldName: string): Promise<string | null> => {
    try {
      const fieldList = await table.getFieldMetaList();
      const target = fieldList.find((f: any) => f.name.trim() === fieldName.trim());
      return target ? target.id : null;
    } catch (e) { return null; }
  };

  // --- LOGICA DE CARGA DE USUARIOS ---
  useEffect(() => {
    async function loadUsers() {
      try {
        const table = await bitable.base.getActiveTable();
        const fieldId = await getFieldIdByNameSafe(table, 'Final User MiRetail (User Code)');
        if (!fieldId) return;
        const records = await table.getRecordIdList();
        const userSet = new Set<string>();
        for (const id of records) {
          const val = await table.getCellValue(fieldId, id);
          const code = parseTextField(val);
          if (code) userSet.add(code);
        }
        setUsers(Array.from(userSet));
      } catch (e) { console.error(e); }
    }
    loadUsers();
  }, []);

  const handleGenerateReport = async (e: React.FormEvent) => {
    e.preventDefault();
    if (!selectedUser) return alert('Selecciona un usuario');
    setLoading(true);
    try {
      const table = await bitable.base.getActiveTable();
      const idUser = await getFieldIdByNameSafe(table, 'Final User MiRetail (User Code)');
      const idName = await getFieldIdByNameSafe(table, 'Final User (Complete Name)');
      const idCat = await getFieldIdByNameSafe(table, 'Category of Expense');
      const idBudget = await getFieldIdByNameSafe(table, 'Estimate Budget');
      const idReleased = await getFieldIdByNameSafe(table, 'Ammount Released (Real)');
      const idInvoice = await getFieldIdByNameSafe(table, 'UUID Invoice');
      const idStart = await getFieldIdByNameSafe(table, 'Fecha de Inicio');
      const idEnd = await getFieldIdByNameSafe(table, 'Fecha de Termino');

      const recordIds = await table.getRecordIdList();
      const filtered: GastoRecord[] = [];
      let fullName = '';
      let sumBudget = 0;
      let sumReleased = 0;
      let minD: number | null = null;
      let maxD: number | null = null;

      const fStart = fechaInicioFiltro ? new Date(fechaInicioFiltro + 'T00:00:00').getTime() : 0;
      const fEnd = fechaTerminoFiltro ? new Date(fechaTerminoFiltro + 'T23:59:59').getTime() : Infinity;

      for (const rid of recordIds) {
        const uVal = await table.getCellValue(idUser!, rid);
        if (parseTextField(uVal) === selectedUser) {
          const nVal = await table.getCellValue(idName!, rid);
          if (!fullName) fullName = parseTextField(nVal);

          const sVal = await table.getCellValue(idStart!, rid);
          const eVal = await table.getCellValue(idEnd!, rid);
          
          const rStart = parseDateField(sVal);
          const rEnd = parseDateField(eVal);

          const passStart = !fechaInicioFiltro || (rStart !== null && rStart >= fStart);
          const passEnd = !fechaTerminoFiltro || (rEnd !== null && rEnd <= fEnd);

          if (passStart && passEnd) {
            const bVal = await table.getCellValue(idBudget!, rid);
            const budget = parseNumberField(bVal);
            sumBudget += budget; 

            const relVal = idReleased ? await table.getCellValue(idReleased, rid) : null;
            const released = parseNumberField(relVal);
            sumReleased += released;

            const invVal = idInvoice ? await table.getCellValue(idInvoice, rid) : null;
            const uuidInvoice = parseTextField(invVal);

            if (rStart !== null && (!minD || rStart < minD)) minD = rStart;
            if (rEnd !== null && (!maxD || rEnd > maxD)) maxD = rEnd;

            filtered.push({
              userCode: selectedUser,
              completeName: fullName,
              category: parseTextField(await table.getCellValue(idCat!, rid)),
              estimateBudget: budget,
              amountReleased: released,
              uuidInvoice,
              fechaInicio: rStart,
              fechaTermino: rEnd
            });
          }
        }
      }

      const diffDays = (minD !== null && maxD !== null) ? Math.ceil((maxD - minD) / (1000 * 60 * 60 * 24)) + 1 : 0;

      setReportData({
        userCode: selectedUser,
        fullName: fullName || selectedUser,
        records: filtered,
        totalBudget: sumBudget,
        totalReleased: sumReleased,
        minDate: minD,
        maxDate: maxD,
        totalDays: diffDays
      });
    } catch (err: any) { alert(err.message); }
    setLoading(false);
  };

  // --- EXPORTACIÓN ESTRUCTURADA CON EXCELJS ---
  useEffect(() => {
    if (!reportData) return;

    const generateExcel = async () => {
      const workbook = new ExcelJS.Workbook();
      const worksheet = workbook.addWorksheet('Comprobación');

      // Configuración de columnas (A hasta G)
      worksheet.columns = [
        { key: 'A', width: 5 },   // #
        { key: 'B', width: 12 },  // FECHA
        { key: 'C', width: 40 },  // FOLIO FISCAL
        { key: 'D', width: 25 },  // TIPO DE GASTO
        { key: 'E', width: 15 },  // MONTO
        { key: 'F', width: 15 },  // TIPO COMPROBANTE
        { key: 'G', width: 20 }   // COMENTARIOS
      ];

      const borderThin = {
        top: { style: 'thin' as const },
        left: { style: 'thin' as const },
        bottom: { style: 'thin' as const },
        right: { style: 'thin' as const }
      };

      // 1. NAME AGENCY
      worksheet.mergeCells('A1:C1');
      const cellAgency = worksheet.getCell('A1');
      cellAgency.value = 'NAME AGENCY';
      cellAgency.font = { name: 'Arial', bold: true, size: 12 };

      // 2. TITULO principal (Barra naranja)
      worksheet.mergeCells('A3:G3');
      const cellTitle = worksheet.getCell('A3');
      cellTitle.value = 'FORMATO DE COMPROBACIÓN';
      cellTitle.font = { name: 'Arial', bold: true, size: 14, color: { argb: 'FFFFFF' } };
      cellTitle.alignment = { horizontal: 'center', vertical: 'middle' };
      cellTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D35400' } };

      // 3. BLOQUE IZQUIERDO (Metadata Usuario)
      const metaLeft = [
        ['USUARIO:', reportData.fullName],
        ['PUESTO:', '-'],
        ['BUDGET AUTORIZADO:', reportData.totalBudget],
        ['TOTAL COMPROBADO:', reportData.totalReleased],
        ['SALDO FINAL:', reportData.totalBudget - reportData.totalReleased],
        ['DESTINO:', '-']
      ];

      metaLeft.forEach((row, index) => {
        const rowIndex = 5 + index;
        worksheet.mergeCells(`B${rowIndex}:C${rowIndex}`);
        
        const labelCell = worksheet.getCell(`A${rowIndex}`);
        labelCell.value = row[0];
        labelCell.font = { name: 'Arial', bold: true, size: 9 };
        labelCell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: row[0] === 'BUDGET AUTORIZADO:' ? 'D4E6F1' : 'EBF5FB' } };
        labelCell.border = borderThin;

        const valCell = worksheet.getCell(`B${rowIndex}`);
        valCell.value = row[1];
        valCell.font = { name: 'Arial', size: 9, bold: typeof row[1] === 'number' };
        valCell.border = borderThin;
        worksheet.getCell(`C${rowIndex}`).border = borderThin; // Aplicar borde al merge

        if (typeof row[1] === 'number') {
          valCell.numFmt = '$#,##0.00';
          if (row[0] === 'SALDO FINAL:') {
            valCell.font = { name: 'Arial', bold: true, size: 9, color: { argb: (row[1] < 0) ? 'D32F2F' : '2E7D32' } };
          }
        }
      });

      // 4. BLOQUE DERECHO (Metadata Fechas y Periodo)
      // Fecha Comprobación
      worksheet.mergeCells('E5:G5');
      const cellFComp = worksheet.getCell('E5');
      cellFComp.value = `FECHA COMPROBACIÓN: ${new Date().toLocaleDateString('es-MX')}`;
      cellFComp.font = { name: 'Arial', size: 9 };
      cellFComp.alignment = { horizontal: 'right' };
      cellFComp.border = borderThin;
      worksheet.getCell('F5').border = borderThin; worksheet.getCell('G5').border = borderThin;

      // Barra Periodo (Naranja)
      worksheet.mergeCells('E7:G7');
      const cellPerTitle = worksheet.getCell('E7');
      cellPerTitle.value = 'PERIODO DEL VIAJE O GASTO';
      cellPerTitle.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFF' } };
      cellPerTitle.alignment = { horizontal: 'center' };
      cellPerTitle.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D35400' } };

      // Fecha Inicio
      worksheet.getCell('E8').value = 'FECHA INICIO:';
      worksheet.getCell('E8').font = { name: 'Arial', bold: true, size: 8 };
      worksheet.getCell('E8').border = borderThin;
      worksheet.mergeCells('F8:G8');
      const cellMinD = worksheet.getCell('F8');
      cellMinD.value = reportData.minDate ? new Date(reportData.minDate).toLocaleDateString('es-MX') : '-';
      cellMinD.font = { name: 'Arial', size: 9 };
      cellMinD.border = borderThin; worksheet.getCell('G8').border = borderThin;

      // Fecha Final
      worksheet.getCell('E9').value = 'FECHA FINAL:';
      worksheet.getCell('E9').font = { name: 'Arial', bold: true, size: 8 };
      worksheet.getCell('E9').border = borderThin;
      worksheet.mergeCells('F9:G9');
      const cellMaxD = worksheet.getCell('F9');
      cellMaxD.value = reportData.maxDate ? new Date(reportData.maxDate).toLocaleDateString('es-MX') : '-';
      cellMaxD.font = { name: 'Arial', size: 9 };
      cellMaxD.border = borderThin; worksheet.getCell('G9').border = borderThin;

      // Total Días
      worksheet.mergeCells('E11:G11');
      const cellDays = worksheet.getCell('E11');
      cellDays.value = `TOTAL DE DIAS DE VIAJE: ${reportData.totalDays}`;
      cellDays.font = { name: 'Arial', bold: true, size: 9 };
      cellDays.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'F2F4F4' } };
      cellDays.border = borderThin;
      worksheet.getCell('F11').border = borderThin; worksheet.getCell('G11').border = borderThin;

      // 5. ENCABEZADOS DE LA TABLA PRINCIPAL (Fila 13)
      const headers = ['#', 'FECHA', 'FOLIO FISCAL / FOLIO DE FACTURA', 'TIPO DE GASTO', 'MONTO', 'TIPO DE COMPROBANTE', 'COMENTARIOS'];
      headers.forEach((h, idx) => {
        const cell = worksheet.getCell(13, idx + 1);
        cell.value = h;
        cell.font = { name: 'Arial', bold: true, size: 9, color: { argb: 'FFFFFF' } };
        cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'D35400' } };
        cell.alignment = { horizontal: 'center', vertical: 'middle', wrapText: true };
        cell.border = borderThin;
      });

      // 6. RENDERIZAR LAS FILAS DE DATOS (Fila 14 en adelante)
      let currentLine = 14;
      const totalRowsToRender = Math.max(20, reportData.records.length);

      for (let i = 0; i < totalRowsToRender; i++) {
        const r = reportData.records[i];
        const rowValues = [
          i + 1,
          r ? (r.fechaInicio ? new Date(r.fechaInicio).toLocaleDateString('es-MX', { day: '2-digit', month: '2-digit' }) : '-') : '',
          r ? r.uuidInvoice : '',
          r ? r.category : '',
          r ? r.amountReleased : '',
          r ? (r.uuidInvoice ? 'INTERNO' : '') : '',
          ''
        ];

        rowValues.forEach((val, idx) => {
          const cell = worksheet.getCell(currentLine, idx + 1);
          cell.value = val;
          cell.font = { name: 'Arial', size: 9 };
          cell.border = borderThin;

          // Alineaciones específicas
          if (idx === 0 || idx === 1 || idx === 5) cell.alignment = { horizontal: 'center' };
          if (idx === 4 && r) { 
            cell.alignment = { horizontal: 'right' }; 
            cell.numFmt = '$#,##0.00'; 
          }

          // Alternancia de color en filas
          if (currentLine % 2 === 0) {
            cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FBFAF9' } };
          }
        });
        currentLine++;
      }

      // 7. FIRMA DE CONFORMIDAD
      currentLine += 2;
      worksheet.mergeCells(`B${currentLine}:E${currentLine}`);
      const cellLineFirma = worksheet.getCell(`B${currentLine}`);
      cellLineFirma.value = '_____________________________________';
      cellLineFirma.alignment = { horizontal: 'center' };

      currentLine++;
      worksheet.mergeCells(`B${currentLine}:E${currentLine}`);
      const cellFirmaText = worksheet.getCell(`B${currentLine}`);
      cellFirmaText.value = `Firma de Conformidad: ${reportData.fullName}`;
      cellFirmaText.font = { name: 'Arial', bold: true, size: 10 };
      cellFirmaText.alignment = { horizontal: 'center' };

      // DESCARGA DEL ARCHIVO .XLSX
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      const link = document.createElement('a');
      link.href = URL.createObjectURL(blob);
      link.download = `Comprobacion_${reportData.fullName.replace(/\s+/g, '_')}.xlsx`;
      link.click();

      setReportData(null);
    };

    generateExcel();
  }, [reportData]);

  return (
    <div style={{ padding: '15px', fontFamily: 'Arial, sans-serif' }}>
      <div>
        <h3>Generador de Formato Agency</h3>
        <form onSubmit={handleGenerateReport} style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
          <select value={selectedUser} onChange={e => setSelectedUser(e.target.value)} required style={{ padding: '8px' }}>
            <option value="">-- Selecciona Usuario --</option>
            {users.map((u, i) => <option key={i} value={u}>{u}</option>)}
          </select>
          <input type="date" value={fechaInicioFiltro} onChange={e => setFechaInicioFiltro(e.target.value)} />
          <input type="date" value={fechaTerminoFiltro} onChange={e => setFechaTerminoFiltro(e.target.value)} />
          <button type="submit" disabled={loading} style={{ padding: '10px', background: '#3370ff', color: '#fff', border: 'none', cursor: 'pointer' }}>
            {loading ? 'Generando...' : 'Descargar Formato Excel'}
          </button>
        </form>
      </div>
    </div>
  );
}
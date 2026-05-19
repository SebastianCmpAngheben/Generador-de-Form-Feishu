import React, { useEffect, useState } from 'react';
import { bitable } from '@lark-base-open/js-sdk';

interface GastoRecord {
  userCode: string;
  completeName: string;
  category: string;
  estimateBudget: number;
  amountReleased: number;
  uuidInvoice: string; // Recibirá el folio generado por tu fórmula de Lark
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
  // Modificado para extraer con máxima compatibilidad strings, objetos de fórmula y arreglos de texto de Lark
  const parseTextField = (cellValue: any): string => {
    if (!cellValue) return '';
    if (typeof cellValue === 'string') return cellValue.trim();
    
    // Si Lark devuelve la fórmula como un arreglo de segmentos de texto (Común en concatenaciones de fórmulas)
    if (Array.isArray(cellValue)) {
      return cellValue.map((seg: any) => {
        if (typeof seg === 'string') return seg;
        return seg?.text || seg?.value || '';
      }).join('').trim();
    }
    
    // Si viene como un objeto directo con propiedad text o value
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
      const idInvoice = await getFieldIdByNameSafe(table, 'UUID Invoice'); // Tu columna calculada (fx)
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

            // Procesamos el valor usando el parseTextField mejorado para Fórmulas de Texto
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
              uuidInvoice, // Guardado de manera limpia
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

  useEffect(() => {
    if (reportData) {
      const timer = setTimeout(() => { window.print(); setReportData(null); }, 800);
      return () => clearTimeout(timer);
    }
  }, [reportData]);

  return (
    <div style={{ padding: '15px', fontFamily: 'Arial, sans-serif' }}>
      <div className="no-print">
        <h3>Generador de Formato Agency</h3>
        <form onSubmit={handleGenerateReport} style={{ display: 'flex', flexDirection: 'column', gap: '10px' }}>
          <select value={selectedUser} onChange={e => setSelectedUser(e.target.value)} required style={{ padding: '8px' }}>
            <option value="">-- Selecciona Usuario --</option>
            {users.map((u, i) => <option key={i} value={u}>{u}</option>)}
          </select>
          <input type="date" value={fechaInicioFiltro} onChange={e => setFechaInicioFiltro(e.target.value)} />
          <input type="date" value={fechaTerminoFiltro} onChange={e => setFechaTerminoFiltro(e.target.value)} />
          <button type="submit" disabled={loading} style={{ padding: '10px', background: '#3370ff', color: '#fff', border: 'none', cursor: 'pointer' }}>
            {loading ? 'Generando...' : 'Descargar Formato'}
          </button>
        </form>
      </div>

      {reportData && (
        <div className="print-only container-form">
          {/* HEADER NAME AGENCY */}
          <div className="agency-header">
            <span className="agency-name">NAME AGENCY</span>
          </div>

          {/* TITULO */}
          <div className="title-bar">
            FORMATO DE COMPROBACIÓN
          </div>

          {/* META DATA GRID */}
          <div className="meta-grid">
            <div className="meta-left">
              <div className="row"><label>USUARIO:</label><span>{reportData.fullName}</span></div>
              <div className="row"><label>PUESTO:</label><span>-</span></div>
              <div className="row highlight-row"><label>BUDGET AUTORIZADO:</label><span>$ {reportData.totalBudget.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span></div>
              <div className="row"><label>TOTAL COMPROBADO:</label><span style={{ fontWeight: 'bold' }}>$ {reportData.totalReleased.toLocaleString('en-US', { minimumFractionDigits: 2 })}</span></div>
              <div className="row"><label>SALDO FINAL:</label><span style={{ fontWeight: 'bold', color: (reportData.totalBudget - reportData.totalReleased) < 0 ? '#d32f2f' : '#2e7d32' }}>$ {(reportData.totalBudget - reportData.totalReleased).toLocaleString('en-US', { minimumFractionDigits: 2 })}</span></div>
              <div className="row"><label>DESTINO:</label><span>-</span></div>
            </div>
            <div className="meta-right">
              <div className="date-box">FECHA COMPROBACIÓN: <strong>{new Date().toLocaleDateString('es-MX')}</strong></div>
              <div className="period-box">
                <div className="period-title">PERIODO DEL VIAJE O GASTO</div>
                <div className="row"><label>FECHA INICIO:</label><span>{reportData.minDate ? new Date(reportData.minDate).toLocaleDateString('es-MX') : '-'}</span></div>
                <div className="row"><label>FECHA FINAL:</label><span>{reportData.maxDate ? new Date(reportData.maxDate).toLocaleDateString('es-MX') : '-'}</span></div>
              </div>
              <div className="days-box">TOTAL DE DIAS DE VIAJE: <strong>{reportData.totalDays}</strong></div>
            </div>
          </div>

          {/* MAIN TABLE */}
          <table className="main-table">
            <thead>
              <tr>
                <th style={{ width: '4%' }}>#</th>
                <th style={{ width: '10%' }}>FECHA</th>
                <th style={{ width: '36%' }}>FOLIO FISCAL / FOLIO DE FACTURA</th>
                <th style={{ width: '18%' }}>TIPO DE GASTO</th>
                <th style={{ width: '12%' }}>MONTO</th>
                <th style={{ width: '10%' }}>TIPO DE COMPROBANTE</th>
                <th style={{ width: '10%' }}>COMENTARIOS</th>
              </tr>
            </thead>
            <tbody>
              {reportData.records.map((r, i) => (
                <tr key={i}>
                  <td style={{ textAlign: 'center' }}>{i + 1}</td>
                  <td style={{ textAlign: 'center' }}>{r.fechaInicio ? new Date(r.fechaInicio).toLocaleDateString('es-MX', { day: '2-digit', month: '2-digit' }) : '-'}</td>
                  {/* Aquí se imprime el folio de la fórmula de manera limpia */}
                  <td style={{ fontSize: '8.5px', paddingLeft: '5px', color: '#333', letterSpacing: '0.2px' }}>{r.uuidInvoice}</td>
                  <td>{r.category}</td>
                  <td style={{ textAlign: 'right', paddingRight: '6px' }}>$ {r.amountReleased.toFixed(2)}</td>
                  <td style={{ textAlign: 'center', fontSize: '8.5px' }}>{r.uuidInvoice ? 'INTERNO' : ''}</td>
                  <td></td>
                </tr>
              ))}
              {Array.from({ length: Math.max(0, 20 - reportData.records.length) }).map((_, i) => (
                <tr key={i + reportData.records.length}>
                  <td style={{ textAlign: 'center' }}>{i + reportData.records.length + 1}</td>
                  <td></td><td></td><td></td><td></td><td></td><td></td>
                </tr>
              ))}
            </tbody>
          </table>

          {/* FOOTER */}
          <div className="footer-signature">
            <div className="signature-box">
               <p style={{ margin: '0 0 2px 0' }}>_____________________________________</p>
               <p style={{ margin: 0 }}><strong>Firma de Conformidad: {reportData.fullName}</strong></p>
            </div>
          </div>
        </div>
      )}

      <style>{`
        @media screen { .print-only { display: none; } }
        @media print {
          @page { size: portrait; margin: 8mm 6mm; }
          body * { visibility: hidden; }
          .print-only, .print-only * { visibility: visible; }
          .print-only { position: absolute; left: 0; top: 0; width: 100%; }
          .no-print { display: none !important; }
        }
        
        .container-form { font-family: 'Segoe UI', Arial, sans-serif; border: 1.5px solid #000; padding: 12px; background-color: #fff; box-sizing: border-box; }
        .agency-header { text-align: left; margin-bottom: 4px; }
        .agency-name { font-weight: bold; font-size: 13px; border-bottom: 1px solid #000; padding-right: 40px; color: #333; }
        
        .title-bar { background-color: #D35400; color: white; text-align: center; font-weight: bold; padding: 6px; font-size: 16px; border: 1px solid #000; margin-bottom: 8px; letter-spacing: 0.5px; }
        
        .meta-grid { display: flex; justify-content: space-between; margin-bottom: 8px; gap: 10px; }
        .meta-left { width: 55%; border: 1px solid #000; }
        .meta-right { width: 43%; display: flex; flex-direction: column; gap: 4px; }
        
        .row { display: flex; border-bottom: 1px solid #000; align-items: center; min-height: 20px; }
        .row:last-child { border-bottom: none; }
        .row label { width: 150px; background-color: #EBF5FB; font-size: 10px; padding: 3px 5px; font-weight: bold; border-right: 1px solid #000; display: flex; align-items: center; min-height: 20px; box-sizing: border-box; }
        .row span { padding-left: 8px; font-size: 10px; flex: 1; color: #111; }
        .highlight-row label { background-color: #D4E6F1; }
        .highlight-row span { font-weight: bold; color: #1B4F72; }

        .date-box, .days-box { border: 1px solid #000; padding: 3px 6px; font-size: 10px; background-color: #F2F4F4; box-sizing: border-box; }
        .period-box { border: 1px solid #000; }
        .period-title { background-color: #D35400; color: white; text-align: center; font-size: 9px; font-weight: bold; padding: 2px; }
        .period-box .row label { width: 100px; }

        .main-table { width: 100%; border-collapse: collapse; border: 1px solid #000; margin-top: 4px; table-layout: fixed; }
        .main-table th { background-color: #D35400; color: white; border: 1px solid #000; font-size: 9px; padding: 5px 2px; text-align: center; font-weight: bold; word-wrap: break-word; }
        .main-table td { border: 1px solid #000; height: 19px; font-size: 9px; padding: 1px 4px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
        .main-table tbody tr:nth-of-type(even) { background-color: #FBFAF9; }
        
        .footer-signature { margin-top: 15px; display: flex; justify-content: flex-start; padding-left: 10px; }
        .signature-box { text-align: center; font-size: 11px; color: #222; }
      `}</style>
    </div>
  );
}
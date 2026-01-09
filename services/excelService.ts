
import * as XLSX from 'xlsx';
import { SaleRow, ComparisonDiff, ComparisonResult } from '../types';

export const parseExcelFile = async (file: File): Promise<SaleRow[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target?.result as ArrayBuffer);
        // 使用 cellNF: true 確保格式化資訊被保留
        const workbook = XLSX.read(data, { type: 'array', cellNF: true, cellText: true });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        // 先讀取原始矩陣來尋找標題列
        const rawRows = XLSX.utils.sheet_to_json(worksheet, { header: 1 }) as any[][];
        
        let headerIdx = -1;
        for (let i = 0; i < rawRows.length; i++) {
          if (rawRows[i] && rawRows[i].length > 0 && String(rawRows[i][0]).trim() === '編號') {
            headerIdx = i;
            break;
          }
        }

        if (headerIdx === -1) {
          reject(new Error("找不到 '編號' 欄位，請確認第一欄標題是否為 '編號'。"));
          return;
        }

        // 使用 raw: false 讀取，這會抓取你在 Excel 中看到的格式化內容 (例如日期會是 2026/1/1 而非 46023)
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { 
          range: headerIdx,
          raw: false, 
          defval: "" // 讓空值預設為空字串，方便比對
        }) as SaleRow[];

        // 過濾掉全空的列 (避免讀到 Excel 底部的空白列)
        const filteredData = jsonData.filter(row => 
          Object.values(row).some(val => val !== "" && val !== null && val !== undefined)
        );

        resolve(filteredData);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

const createKey = (row: SaleRow): string => {
  const idVal = String(row['編號'] || '').trim();
  const projVal = String(row['案名'] || '').trim();
  const custVal = String(row['客戶'] || '').trim();
  return `${idVal}_${projVal}_${custVal}`;
};

export const compareData = (oldData: SaleRow[], newData: SaleRow[]): ComparisonResult => {
  const oldMap = new Map<string, SaleRow>();
  const newMap = new Map<string, SaleRow>();

  oldData.forEach(row => {
    const key = createKey(row);
    if (key !== "__") oldMap.set(key, row);
  });
  
  newData.forEach(row => {
    const key = createKey(row);
    if (key !== "__") newMap.set(key, row);
  });

  const targetCols = ['日', '狀態', '產品', 'Qty pc', 'Qty kW'];
  const modified: ComparisonDiff[] = [];
  const added: SaleRow[] = [];
  const removed: SaleRow[] = [];

  // 檢查修改與移除
  oldMap.forEach((oldRow, key) => {
    if (newMap.has(key)) {
      const newRow = newMap.get(key)!;
      const changes: ComparisonDiff['changes'] = [];

      targetCols.forEach(col => {
        const vOld = oldRow[col];
        const vNew = newRow[col];

        const sOld = String(vOld ?? '').trim();
        const sNew = String(vNew ?? '').trim();

        if (sOld !== sNew) {
          changes.push({
            column: col,
            oldValue: sOld || '(空)',
            newValue: sNew || '(空)'
          });
        }
      });

      if (changes.length > 0) {
        modified.push({
          uniqueKey: key,
          id: String(newRow['編號'] || ''),
          projectName: String(newRow['案名'] || ''),
          customer: String(newRow['客戶'] || ''),
          changes
        });
      }
    } else {
      removed.push(oldRow);
    }
  });

  // 檢查新增
  newMap.forEach((newRow, key) => {
    if (!oldMap.has(key)) {
      added.push(newRow);
    }
  });

  return { added, removed, modified };
};

export const exportReport = (result: ComparisonResult) => {
  const reportData = result.modified.map(diff => ({
    '組合索引 (A+I+J)': diff.uniqueKey,
    '編號': diff.id,
    '案名': diff.projectName,
    '客戶': diff.customer,
    '變更明細': diff.changes.map(c => `[${c.column}]: ${c.oldValue} -> ${c.newValue}`).join('\n')
  }));

  const worksheet = XLSX.utils.json_to_sheet(reportData);
  
  // 設定欄位寬度讓報表更好讀
  worksheet['!cols'] = [
    { wch: 40 }, // 索引
    { wch: 15 }, // 編號
    { wch: 30 }, // 案名
    { wch: 20 }, // 客戶
    { wch: 50 }, // 變更明細
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Comparison Report');
  XLSX.writeFile(workbook, 'comparison_report.xlsx');
};

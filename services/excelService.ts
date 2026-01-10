
import * as XLSX from 'xlsx';
import { SaleRow, ComparisonDiff, ComparisonResult } from '../types';

/**
 * 將 Excel 的日期序號（如 46204）轉換為 YYYY/MM/DD 字串
 */
const formatExcelDate = (val: any): string => {
  if (typeof val === 'number' && val > 30000 && val < 60000) {
    try {
      // Excel 序號起點是 1899/12/30
      const date = new Date((val - 25569) * 86400 * 1000);
      const y = date.getFullYear();
      const m = String(date.getMonth() + 1).padStart(2, '0');
      const d = String(date.getDate()).padStart(2, '0');
      return `${y}/${m}/${d}`;
    } catch (e) {
      return String(val);
    }
  }
  return String(val || "").trim();
};

/**
 * 極限標準化：移除所有空格、換行及隱形字元，並轉為字串。
 * 用於比對兩個值是否「意義上」相同。
 */
const extremeNormalize = (val: any): string => {
  if (val === null || val === undefined || String(val).trim() === "") return "";
  // 如果是日期序號，先轉成日期字串再比對
  let str = typeof val === 'number' ? formatExcelDate(val) : String(val);
  // 移除所有空白字元及特殊不可見字元
  return str.replace(/[\s\r\n\t\uFEFF\xA0\u200b-\u200d]+/g, '').trim();
};

/**
 * 尋找欄位名稱的索引，支援模糊匹配（忽略空白、大小寫）
 */
const findColIndex = (headers: string[], targetName: string): number => {
  const normalizedTarget = extremeNormalize(targetName);
  return headers.findIndex(h => extremeNormalize(h).includes(normalizedTarget));
};

export const parseExcelFile = async (file: File): Promise<SaleRow[]> => {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = (e) => {
      try {
        const fileBuffer = new Uint8Array(e.target?.result as ArrayBuffer);
        const workbook = XLSX.read(fileBuffer, { type: 'array' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        
        const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: "" }) as any[][];
        
        let headerIdx = -1;
        for (let i = 0; i < Math.min(rows.length, 100); i++) {
          if (rows[i].some(cell => extremeNormalize(cell).includes("編號"))) {
            headerIdx = i;
            break;
          }
        }

        if (headerIdx === -1) {
          reject(new Error("找不到 '編號' 標題列，請確認 Excel 格式。"));
          return;
        }

        const headers = rows[headerIdx].map(h => String(h));
        
        const colMap = {
          id: findColIndex(headers, "編號"),
          date: findColIndex(headers, "日"),
          status: findColIndex(headers, "狀態"),
          product: findColIndex(headers, "產品"),
          qtyPc: findColIndex(headers, "Qtypc"),
          qtyKw: findColIndex(headers, "QtykW"),
          project: findColIndex(headers, "案名"),
          customer: findColIndex(headers, "客戶")
        };

        const parsedRows: SaleRow[] = [];
        for (let i = headerIdx + 1; i < rows.length; i++) {
          const row = rows[i];
          const idValue = row[colMap.id];
          
          // 如果編號是日期序號，進行格式化
          const formattedId = typeof idValue === 'number' ? formatExcelDate(idValue) : String(idValue || "").trim();
          
          if (!formattedId) continue;

          // 預先處理「日」欄位的格式
          if (colMap.date !== -1 && typeof row[colMap.date] === 'number') {
            row[colMap.date] = formatExcelDate(row[colMap.date]);
          }

          parsedRows.push({
            _raw: row,
            _colMap: colMap,
            id: formattedId,
            projectName: String(row[colMap.project] || "").trim(),
            customer: String(row[colMap.customer] || "").trim()
          });
        }

        resolve(parsedRows);
      } catch (err) {
        reject(err);
      }
    };
    reader.onerror = reject;
    reader.readAsArrayBuffer(file);
  });
};

export const compareData = (oldData: SaleRow[], newData: SaleRow[]): ComparisonResult => {
  const modified: ComparisonDiff[] = [];
  const added: SaleRow[] = [];
  const removed: SaleRow[] = [];

  const compareFields = [
    { label: '日', mapKey: 'date' },
    { label: '狀態', mapKey: 'status' },
    { label: '產品', mapKey: 'product' },
    { label: 'Qty pc', mapKey: 'qtyPc' },
    { label: 'Qty kW', mapKey: 'qtyKw' }
  ];

  const oldMap = new Map<string, SaleRow>();
  oldData.forEach(row => {
    const idKey = extremeNormalize(row.id);
    if (idKey) oldMap.set(idKey, row);
  });

  const processedIds = new Set<string>();

  newData.forEach(newRow => {
    const idKey = extremeNormalize(newRow.id);
    if (!idKey) return;
    processedIds.add(idKey);

    if (oldMap.has(idKey)) {
      const oldRow = oldMap.get(idKey)!;
      const changes: ComparisonDiff['changes'] = [];

      compareFields.forEach(field => {
        const oldIdx = oldRow._colMap[field.mapKey];
        const newIdx = newRow._colMap[field.mapKey];
        
        if (oldIdx !== -1 && newIdx !== -1) {
          const valOld = extremeNormalize(oldRow._raw[oldIdx]);
          const valNew = extremeNormalize(newRow._raw[newIdx]);

          if (valOld !== valNew) {
            changes.push({
              column: field.label,
              oldValue: valOld || "(空)",
              newValue: valNew || "(空)"
            });
          }
        }
      });

      if (changes.length > 0) {
        modified.push({
          uniqueKey: idKey,
          id: newRow.id,
          projectName: newRow.projectName,
          customer: newRow.customer,
          changes
        });
      }
    } else {
      added.push(newRow);
    }
  });

  oldMap.forEach((row, id) => {
    if (!processedIds.has(id)) {
      removed.push(row);
    }
  });

  return { added, removed, modified };
};

export const exportReport = (result: ComparisonResult) => {
  const finalData: any[] = [];

  // 1. 處理異動項次
  result.modified.forEach(diff => {
    finalData.push({
      '狀態類型': '內容異動',
      '編號': diff.id,
      '案名': diff.projectName,
      '客戶': diff.customer,
      '變更明細': diff.changes.map(c => `【${c.column}】 ${c.oldValue} -> ${c.newValue}`).join(' ; ')
    });
  });

  // 2. 處理全新增項次
  result.added.forEach(row => {
    const details = [
      row._colMap.date !== -1 ? `[日: ${row._raw[row._colMap.date]}]` : '',
      row._colMap.status !== -1 ? `[狀態: ${row._raw[row._colMap.status]}]` : '',
      row._colMap.product !== -1 ? `[產品: ${row._raw[row._colMap.product]}]` : '',
      row._colMap.qtyPc !== -1 ? `[Qty pc: ${row._raw[row._colMap.qtyPc]}]` : '',
    ].filter(Boolean).join(' ');

    finalData.push({
      '狀態類型': '全新增項',
      '編號': row.id,
      '案名': row.projectName,
      '客戶': row.customer,
      '變更明細': `新增資料：${details}`
    });
  });

  // 3. 處理已移除項次
  result.removed.forEach(row => {
    finalData.push({
      '狀態類型': '已移除項',
      '編號': row.id,
      '案名': row.projectName,
      '客戶': row.customer,
      '變更明細': '此項次已從更新檔案中刪除'
    });
  });

  const worksheet = XLSX.utils.json_to_sheet(finalData);
  
  // 設定欄寬
  worksheet['!cols'] = [
    { wch: 12 }, // 狀態類型
    { wch: 15 }, // 編號
    { wch: 40 }, // 案名
    { wch: 20 }, // 客戶
    { wch: 80 }  // 變更明細
  ];

  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, worksheet, '業務異動報告');
  XLSX.writeFile(workbook, `業務比對完整報告_${new Date().toISOString().split('T')[0]}.xlsx`);
};

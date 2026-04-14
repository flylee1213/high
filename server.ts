import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import { fileURLToPath } from "url";
import multer from "multer";
import * as XLSX from "xlsx";
import PizZip from "pizzip";
import Docxtemplater from "docxtemplater";
import ExcelJS from "exceljs";
import AdmZip from "adm-zip";
import fs from "fs/promises";
import { existsSync } from "fs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

function escapeRegExp(string: string) {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

// Mapping Configuration: Excel Column Name -> Word Keyword
const MAPPING_CONFIG: Record<string, string> = {
  "单位": "[单位]",
  "合同编号": "[合同编号]",
  "日期": "[DATE]",
  "金额": "[AMOUNT]",
  "甲方": "[PARTY_A]",
  "乙方": "[PARTY_B]"
};

// Brand and Model Mapping Table
const BRAND_MODEL_MAPPING: Record<string, Record<string, string>> = {
  "台式机": {
    "华为": "华为擎云W515x",
    "国光": "UT6000-KB1",
    "联想": "开天E50h G1t",
    "同方": "超翔H880-L52"
  },
  "笔记本": {
    "华为": "华为擎云L420x",
    "国光": "UT6500-LA6",
    "联想": "开天N60d G1d",
    "同方": "超锐L60P-13031"
  }
};

async function startServer() {
  const app = express();
  const PORT = 3000;

  // Configure Multer for file uploads
  const upload = multer({ storage: multer.memoryStorage() });

  // API Route: Get Mapping Config for UI
  app.get("/api/mapping", (req, res) => {
    res.json(MAPPING_CONFIG);
  });

  // API Route: Upload and Process
  app.post("/api/process", upload.fields([
    { name: 'dataSource', maxCount: 1 },
    { name: 'templates', maxCount: 4 }
  ]), async (req: any, res) => {
    try {
      const files = req.files;
      if (!files.dataSource || !files.templates) {
        return res.status(400).json({ error: "Missing data source or templates" });
      }

      // Fix filename encoding
      const fixEncoding = (filename: string) => Buffer.from(filename, 'latin1').toString('utf8');
      const dataSourceFile = files.dataSource[0];
      const templateFiles = files.templates.map((f: any) => {
        f.originalname = fixEncoding(f.originalname);
        return f;
      });

      // 1. Parse Data Source with XLSX
      const workbook = XLSX.read(dataSourceFile.buffer, { type: 'buffer' });
      
      const zip = new AdmZip();
      const timestamp = Date.now();

      // Iterate through all sheets (Regions)
      for (const sheetName of workbook.SheetNames) {
        const rawData = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { defval: "" });
        if (rawData.length === 0) continue;

        // Grouping Logic per sheet: Fill down and group by "单位" or "合同编号"
        const groups: Record<string, any[]> = {};
        let lastMainInfo: any = {};
        // Add brands to mainKeys to ensure they are filled down within a group
        const mainKeys = ["单位", "合同编号", "序号", "甲方", "乙方", "日期", "金额", "台式机品牌", "笔记本品牌"];

        (rawData as any[]).forEach((row) => {
          const hasMainInfo = row["单位"] || row["序号"] || row["合同编号"];
          
          if (hasMainInfo) {
            lastMainInfo = {};
            mainKeys.forEach(key => {
              if (row[key]) lastMainInfo[key] = row[key];
            });
          }

          const fullRow = { ...row };
          mainKeys.forEach(key => {
            if (!fullRow[key] && lastMainInfo[key]) {
              fullRow[key] = lastMainInfo[key];
            }
          });

          const groupKey = `${fullRow["单位"] || ""}_${fullRow["合同编号"] || ""}`;
          if (!groups[groupKey]) {
            groups[groupKey] = [];
          }
          groups[groupKey].push(fullRow);
        });

        // 2. Process each group in the current sheet
        for (const groupRows of Object.values(groups)) {
          const firstRow = groupRows[0];
          
          // Calculate sums for numeric columns within the group
          const groupTotals: Record<string, number> = {};
          const numericKeys = new Set<string>();

          groupRows.forEach(row => {
            Object.entries(row).forEach(([key, val]) => {
              if (val !== "" && !isNaN(Number(val)) && typeof val !== 'boolean') {
                numericKeys.add(key);
                groupTotals[key] = (groupTotals[key] || 0) + Number(val);
              }
            });
          });

          // Module 3: Acceptance Items Logic
          const acceptanceItems: any[] = [];
          let currentId = 1;

          const collectItems = (type: "台式机" | "笔记本") => {
            const brandMap: Record<string, number> = {};
            groupRows.forEach(row => {
              const brand = (row[`${type}品牌`] || "").toString();
              const qty = Number(row[`${type}数量`]) || 0;
              if (brand && qty > 0) {
                brandMap[brand] = (brandMap[brand] || 0) + qty;
              }
            });

            Object.entries(brandMap).forEach(([brand, qty]) => {
              const item = {
                "序号": currentId,
                "id": currentId,
                "设备名称": type,
                "deviceName": type,
                "品牌": brand,
                "brand": brand,
                "型号": BRAND_MODEL_MAPPING[type]?.[brand] || "",
                "model": BRAND_MODEL_MAPPING[type]?.[brand] || "",
                "数量": qty,
                "quantity": qty
              };
              acceptanceItems.push(item);
              currentId++;
            });
          };

          collectItems("台式机");
          collectItems("笔记本");

          // Module 4: Device List Expansion Logic
          const deviceList: any[] = [];
          let deviceId = 1;

          groupRows.forEach(row => {
            const types: ("台式机" | "笔记本")[] = ["台式机", "笔记本"];
            types.forEach(type => {
              const brand = (row[`${type}品牌`] || "").toString();
              const os = (row[`${type}OS`] || "").toString();
              const qty = Number(row[`${type}数量`]) || 0;
              
              if (brand && qty > 0) {
                for (let i = 0; i < qty; i++) {
                  deviceList.push({
                    "序号": deviceId,
                    "id": deviceId,
                    "设备类型": type,
                    "终端品牌": brand,
                    "品牌": brand, // Add alias to match [品牌] in template
                    "设备型号": row[`${type}型号`] || BRAND_MODEL_MAPPING[type]?.[brand] || "",
                    "终端SN串号": "",
                    "操作系统": os
                  });
                  deviceId++;
                }
              }
            });
          });

          const hasDesktop = acceptanceItems.some(i => i.deviceName === "台式机");
          const hasLaptop = acceptanceItems.some(i => i.deviceName === "笔记本");

          // Helper to get value (sum if numeric, otherwise first row's value)
          const getVal = (key: string) => {
            if (numericKeys.has(key)) {
              return groupTotals[key].toString();
            }
            
            // Special Logic: Automatic Model Lookup
            if (key === "型号") {
              const type = (firstRow["类型"] || "").toString();
              const brand = (firstRow["品牌"] || "").toString();
              if (BRAND_MODEL_MAPPING[type]?.[brand]) {
                return BRAND_MODEL_MAPPING[type][brand];
              }
              return "";
            }
            if (key === "台式机型号") {
              const brand = (firstRow["台式机品牌"] || "").toString();
              return BRAND_MODEL_MAPPING["台式机"]?.[brand] || "";
            }
            if (key === "笔记本型号") {
              const brand = (firstRow["笔记本品牌"] || "").toString();
              return BRAND_MODEL_MAPPING["笔记本"]?.[brand] || "";
            }

            return (firstRow[key] || "").toString();
          };

          // Prepare data for replacement
          const replacementData: Record<string, any> = {
            items: groupRows,
            acceptanceItems: acceptanceItems,
            deviceList: deviceList,
            hasDesktop: hasDesktop,
            hasLaptop: hasLaptop,
            "【台式机】": hasDesktop ? "台式机" : "",
            "【笔记本】": hasLaptop ? "笔记本" : ""
          };

          // Add top-level keys for simple replacement based on MAPPING_CONFIG
          Object.entries(MAPPING_CONFIG).forEach(([excelCol, wordKeyword]) => {
            replacementData[wordKeyword] = getVal(excelCol);
          });

          // Add specific auto-lookup fields to replacement data
          replacementData["[型号]"] = getVal("型号");
          replacementData["[台式机型号]"] = BRAND_MODEL_MAPPING["台式机"]?.[firstRow["台式机品牌"]] || "";
          replacementData["[笔记本型号]"] = BRAND_MODEL_MAPPING["笔记本"]?.[firstRow["笔记本品牌"]] || "";

          const rowIdentifier = firstRow["单位"] || `Contract_${Math.random().toString(36).slice(2, 7)}`;
          const safeRowName = rowIdentifier.toString().replace(/[/\\?%*:|"<>]/g, '-');
          const safeSheetName = sheetName.toString().replace(/[/\\?%*:|"<>]/g, '-');

          for (const template of templateFiles) {
            const extension = path.extname(template.originalname).toLowerCase();
            const baseName = path.basename(template.originalname, extension);
            const outputFileName = `${safeRowName}${baseName}${extension}`;

            if (extension === '.docx') {
              try {
                if (template.buffer[0] !== 0x50 || template.buffer[1] !== 0x4B) {
                  throw new Error(`文件 "${template.originalname}" 不是有效的 .docx 文件。`);
                }

                const zipDoc = new PizZip(template.buffer);
                const doc = new Docxtemplater(zipDoc, {
                  paragraphLoop: true,
                  linebreaks: true,
                  delimiters: { start: "[", end: "]" }
                });

                const docData: Record<string, any> = { items: groupRows, deviceList: deviceList };
                Object.keys(firstRow).forEach(key => { docData[key] = getVal(key); });
                Object.entries(replacementData).forEach(([keyword, value]) => {
                  if (keyword !== 'items') {
                    const cleanKey = keyword.replace(/[\[\]]/g, "");
                    docData[cleanKey] = value;
                  }
                });

                if (acceptanceItems.length > 0) {
                  Object.entries(acceptanceItems[0]).forEach(([k, v]) => {
                    if (docData[k] === undefined || docData[k] === "") docData[k] = v;
                  });
                }

                doc.render(docData);
                const buf = doc.getZip().generate({ type: "nodebuffer" });
                zip.addFile(`${safeSheetName}/${safeRowName}/${outputFileName}`, buf);
              } catch (err: any) {
                console.error(`Error processing ${template.originalname}:`, err);
                throw new Error(`Failed to process "${template.originalname}": ${err.message}`);
              }
            } else if (extension === '.xlsx') {
              try {
                if (template.buffer[0] !== 0x50 || template.buffer[1] !== 0x4B) {
                  throw new Error(`文件 "${template.originalname}" 不是有效的 .xlsx 文件。`);
                }

                const templateWorkbook = new ExcelJS.Workbook();
                await templateWorkbook.xlsx.load(template.buffer);

                templateWorkbook.eachSheet((sheet) => {
                  const rowsToLoop: { rowNumber: number, dataKey: string }[] = [];
                  
                  // 1. Identify rows that contain loop tags [#key] ... [/key]
                  sheet.eachRow((excelRow, rowNumber) => {
                    let hasStart = false;
                    let hasEnd = false;
                    let key = "items";
                    
                    excelRow.eachCell((cell) => {
                      if (typeof cell.value === 'string') {
                        const startMatch = cell.value.match(/\[#(\w+)\]/);
                        if (startMatch) {
                          hasStart = true;
                          key = startMatch[1];
                        }
                        if (cell.value.includes(`[/${key}]`)) hasEnd = true;
                      }
                    });
                    
                    if (hasStart && hasEnd) {
                      rowsToLoop.push({ rowNumber, dataKey: key });
                    }
                  });

                  // 2. Process loops (in reverse to not mess up row numbers)
                  rowsToLoop.sort((a, b) => b.rowNumber - a.rowNumber).forEach(loopInfo => {
                    const templateRow = sheet.getRow(loopInfo.rowNumber);
                    const dataList = replacementData[loopInfo.dataKey] || [];
                    
                    if (!Array.isArray(dataList) || dataList.length === 0) {
                      // If no data, we can either clear the row or delete it. 
                      // For "Device List", usually we delete the template row if empty.
                      sheet.spliceRows(loopInfo.rowNumber, 1);
                      return;
                    }

                    // Insert necessary rows (pass an array of empty arrays to create empty rows)
                    if (dataList.length > 1) {
                      sheet.insertRows(loopInfo.rowNumber + 1, new Array(dataList.length - 1).fill([]));
                    }

                    // Fill data
                    const templateCells: { col: number, value: any }[] = [];
                    templateRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                      templateCells.push({ col: colNumber, value: cell.value });
                    });

                    dataList.forEach((itemData, index) => {
                      const currentRow = sheet.getRow(loopInfo.rowNumber + index);
                      
                      // Copy styles from template row
                      if (index > 0) {
                        currentRow.height = templateRow.height;
                        templateRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
                          const newCell = currentRow.getCell(colNumber);
                          newCell.style = { ...cell.style };
                        });
                      }

                      templateCells.forEach(({ col, value }) => {
                        if (typeof value === 'string') {
                          let cellValue = value;
                          
                          // Remove loop tags
                          cellValue = cellValue.replace(/\[#\w+\]/g, '').replace(/\[\/\w+\]/g, '');
                          
                          // Prepare row-specific data
                          const rowData = { 
                            ...firstRow, 
                            ...itemData,
                            "序号": index + 1,
                            "id": index + 1 
                          };

                          // Replace placeholders
                          Object.entries(rowData).forEach(([k, v]) => {
                            const placeholder = `[${k}]`;
                            if (cellValue.includes(placeholder)) {
                              const escaped = escapeRegExp(placeholder);
                              cellValue = cellValue.replace(new RegExp(escaped, 'g'), (v || "").toString());
                            }
                          });

                          // Also support MAPPING_CONFIG keywords inside loops
                          Object.entries(MAPPING_CONFIG).forEach(([excelCol, wordKeyword]) => {
                            if (cellValue.includes(wordKeyword)) {
                              const val = rowData[excelCol] || "";
                              const escaped = escapeRegExp(wordKeyword);
                              cellValue = cellValue.replace(new RegExp(escaped, 'g'), val.toString());
                            }
                          });

                          currentRow.getCell(col).value = cellValue;
                        } else {
                          currentRow.getCell(col).value = value;
                        }
                      });
                    });
                  });

                  // 3. Regular non-loop replacement for the rest of the sheet
                  sheet.eachRow((excelRow) => {
                    // Skip rows that were part of a loop (they are already processed)
                    // Simple check: if it still has brackets, it might need replacement
                    excelRow.eachCell((cell) => {
                      if (typeof cell.value === 'string' && cell.value.includes("[")) {
                        let cellValue = cell.value;
                        
                        // Standard replacements
                        Object.entries(replacementData).forEach(([keyword, val]) => {
                          if (keyword !== 'items' && keyword !== 'acceptanceItems' && cellValue.includes(keyword)) {
                            const escapedPlaceholder = escapeRegExp(keyword);
                            cellValue = cellValue.replace(new RegExp(escapedPlaceholder, 'g'), val.toString());
                          }
                        });

                        Object.keys(firstRow).forEach((key) => {
                          const placeholder = `[${key}]`;
                          if (cellValue.includes(placeholder)) {
                            const val = getVal(key);
                            const escapedPlaceholder = escapeRegExp(placeholder);
                            cellValue = cellValue.replace(new RegExp(escapedPlaceholder, 'g'), val);
                          }
                        });

                        // Smart Cleanup & Safety Net
                        if (cell.value !== cellValue || cellValue.includes("[")) {
                          cellValue = cellValue.replace(/\[[^\]]+\]/g, '');
                          cellValue = cellValue.replace(/^[、，,；; \-]+|[、，,；; \-]+$/g, '');
                          cellValue = cellValue.replace(/[、，,；; \-]{2,}/g, (match) => match[0]);
                          if (/^[、，,；; \-]*$/.test(cellValue)) cellValue = "";
                        }

                        cell.value = cellValue;
                      }
                    });
                  });
                });

                const buf = await templateWorkbook.xlsx.writeBuffer();
                zip.addFile(`${safeSheetName}/${safeRowName}/${outputFileName}`, Buffer.from(buf));
              } catch (err: any) {
                console.error(`Error processing ${template.originalname}:`, err);
                throw new Error(`Failed to process "${template.originalname}": ${err.message}`);
              }
            }
          }
        }
      }

      const zipBuffer = zip.toBuffer();

      res.writeHead(200, {
        'Content-Type': 'application/zip',
        'Content-Disposition': `attachment; filename="contracts_${timestamp}.zip"`,
        'Content-Length': zipBuffer.length
      });

      res.end(zipBuffer);

    } catch (error: any) {
      console.error("Processing error:", error);
      res.status(500).json({ error: error.message || "Internal server error" });
    }
  });

  // Vite middleware for development
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();

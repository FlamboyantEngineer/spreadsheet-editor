import type ExcelJS from 'exceljs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';

/** ExcelJS only writes fullCalcOnLoad. Store manual calculation in the workbook part. */
export async function writeManualWorkbook(document: ExcelJS.Workbook): Promise<Uint8Array> {
  const files = unzipSync(new Uint8Array(await document.xlsx.writeBuffer()));
  const xml = strFromU8(files['xl/workbook.xml']);
  files['xl/workbook.xml'] = strToU8(
    xml.replace(
      /<calcPr\b[^>]*\/?\s*>/,
      '<calcPr calcMode="manual" fullCalcOnLoad="0" forceFullCalc="0" calcOnSave="0"/>',
    ),
  );
  return zipSync(files);
}

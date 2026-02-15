import {
  AlignmentType,
  BorderStyle,
  convertInchesToTwip,
  Document,
  Packer,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from "docx";
import * as fs from "fs";
import * as path from "path";
import * as puppeteer from "puppeteer";

export interface ReportStats {
  total: number;
  women: number;
  men: number;
  unemployed: number;
  pensioners: number;
  disabled: number;
}

/**
 * Format date as dd.mm.yyyy for report title
 */
function formatReportDate(date: Date): string {
  const day = String(date.getDate()).padStart(2, "0");
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const year = date.getFullYear();
  return `${day}.${month}.${year}`;
}

/** Shared timestamp for report filenames (Word and PDF same base) */
export function getReportFileTimestamp(): string {
  return new Date()
    .toLocaleString("uz-UZ", {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
    })
    .replace(/[/\s:,]+/g, "-");
}

/**
 * Generates a Word document for the patient summary report (from-to date range, table: Jami, Ayollar, Erkaklar, Ishsizlar, Nafaqaxo'rlar, Nogironlar).
 * If from/to are not provided, dateFrom/dateTo from stats (min/max createdAt) are used for the title.
 */
export async function generateReportWord(
  stats: ReportStats,
  dateFrom: Date,
  dateTo: Date,
  fileTimestamp?: string,
): Promise<string> {
  const dateFromStr = formatReportDate(dateFrom);
  const dateToStr = formatReportDate(dateTo);
  const title = `${dateFromStr} - ${dateToStr} oraliqda kelgan patsientlar bo'yicha hisobot`;

  const dateTime =
    fileTimestamp ??
    new Date()
      .toLocaleString("uz-UZ", {
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
      })
      .replace(/[/\s:,]+/g, "-");

  const doc = new Document({
    sections: [
      {
        properties: {
          page: {
            margin: {
              top: convertInchesToTwip(0.5),
              right: convertInchesToTwip(0.5),
              bottom: convertInchesToTwip(0.5),
              left: convertInchesToTwip(0.5),
            },
          },
        },
        children: [
          new Paragraph({
            children: [
              new TextRun({
                text: title,
                bold: true,
                size: 28,
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 400 },
          }),
          createSummaryTable(stats),
          new Paragraph({ children: [], spacing: { after: 500 } }),
          new Paragraph({
            children: [
              new TextRun({
                text: "Diqqat!!!",
                bold: true,
                color: "FF0000",
                size: 28,
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
          }),
          new Paragraph({
            children: [
              new TextRun({
                text:
                  "Ushbu ko'rinishda oddiy jadval tuzilib, tizimda hisobot yaratish qanday ishlashi taxminan ko'rsatilmoqda. Albatta, har bir mijozga tizim orqali o'ziga xohishga mos hisobotlarni yaratish imkoniyati beriladi. Bu faqat namuna hisobot shakli.",
                size: 24,
                bold: true,
              }),
            ],
            alignment: AlignmentType.CENTER,
            spacing: { after: 200 },
          }),
        ],
      },
    ],
  });

  const fileName = `hisobot-${dateTime}.docx`;
  const filePath = path.join(process.cwd(), "uploads", fileName);

  const uploadsDir = path.join(process.cwd(), "uploads");
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
  }

  const buffer = await Packer.toBuffer(doc);
  fs.writeFileSync(filePath, buffer);

  return `uploads/${fileName}`;
}

/**
 * Generates a PDF with the same content as the report Word (title, summary table, Diqqat!!!, disclaimer).
 */
export async function generateReportPdf(
  stats: ReportStats,
  dateFrom: Date,
  dateTo: Date,
  fileTimestamp?: string,
): Promise<string> {
  const dateTime =
    fileTimestamp ??
    new Date()
      .toLocaleString("uz-UZ", {
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
        second: "2-digit",
      })
      .replace(/[/\s:,]+/g, "-");

  const html = buildReportHtml(stats, dateFrom, dateTo);

  const browser = await puppeteer.launch({
    headless: true,
    args: [
      "--no-sandbox",
      "--disable-setuid-sandbox",
      "--disable-dev-shm-usage",
      "--disable-accelerated-2d-canvas",
      "--no-first-run",
      "--no-zygote",
      "--single-process",
      "--disable-gpu",
    ],
  });

  const page = await browser.newPage();
  await page.setContent(html, { waitUntil: "networkidle0" });

  const fileName = `hisobot-${dateTime}.pdf`;
  const filePath = path.join(process.cwd(), "uploads", fileName);

  const uploadsDir = path.join(process.cwd(), "uploads");
  if (!fs.existsSync(uploadsDir)) {
    fs.mkdirSync(uploadsDir, { recursive: true });
  }

  await page.pdf({
    path: filePath,
    format: "A4",
    margin: {
      top: "0.5in",
      right: "0.5in",
      bottom: "0.5in",
      left: "0.5in",
    },
    printBackground: true,
  });

  await browser.close();

  return `uploads/${fileName}`;
}

function buildReportHtml(
  stats: ReportStats,
  dateFrom: Date,
  dateTo: Date,
): string {
  const dateFromStr = formatReportDate(dateFrom);
  const dateToStr = formatReportDate(dateTo);
  const title = `${dateFromStr} - ${dateToStr} oraliqda kelgan patsientlar bo'yicha hisobot`;

  return `
<!DOCTYPE html>
<html>
<head>
  <meta charset="UTF-8">
  <style>
    body { font-family: 'Times New Roman', Arial, sans-serif; margin: 0; padding: 0; }
    .title { text-align: center; font-weight: bold; font-size: 18pt; margin-bottom: 24px; }
    .diqqat { text-align: center; font-weight: bold; color: #c00; font-size: 14pt; margin: 24px 0 8px; }
    .disclaimer { text-align: center; font-size: 11pt; margin-bottom: 16px; line-height: 1.4; font-weight: bold; }
    table { width: 100%; border-collapse: collapse; margin-bottom: 28px; }
    th, td { border: 1px solid #000; padding: 10px; text-align: center; font-size: 12pt; }
    th { font-weight: bold; }
  </style>
</head>
<body>
  <h1 class="title">${title}</h1>
  <table>
    <thead><tr>
      <th>Jami</th><th>Ayollar</th><th>Erkaklar</th><th>Ishsizlar</th><th>Nafaqaxo'rlar</th><th>Nogironlar</th>
    </tr></thead>
    <tbody><tr>
      <td>${stats.total}</td><td>${stats.women}</td><td>${stats.men}</td>
      <td>${stats.unemployed}</td><td>${stats.pensioners}</td><td>${stats.disabled}</td>
    </tr></tbody>
  </table>
  <p class="diqqat">Diqqat!!!</p>
  <p class="disclaimer">Ushbu ko'rinishda oddiy jadval tuzilib, tizimda hisobot yaratish qanday ishlashi taxminan ko'rsatilmoqda. Albatta, har bir mijozga tizim orqali o'ziga xohishga mos hisobotlarni yaratish imkoniyati beriladi. Bu faqat namuna hisobot shakli.</p>
</body>
</html>`;
}

const COLUMNS = [
  "Jami",
  "Ayollar",
  "Erkaklar",
  "Ishsizlar",
  "Nafaqaxo'rlar",
  "Nogironlar",
] as const;

const COLUMN_WIDTH = 100 / COLUMNS.length;

function createSummaryTable(stats: ReportStats): Table {
  const headerRow = new TableRow({
    children: COLUMNS.map((label) =>
      createReportCell(label, true, COLUMN_WIDTH),
    ),
  });

  const dataRow = new TableRow({
    children: [
      createReportCell(String(stats.total), false, COLUMN_WIDTH),
      createReportCell(String(stats.women), false, COLUMN_WIDTH),
      createReportCell(String(stats.men), false, COLUMN_WIDTH),
      createReportCell(String(stats.unemployed), false, COLUMN_WIDTH),
      createReportCell(String(stats.pensioners), false, COLUMN_WIDTH),
      createReportCell(String(stats.disabled), false, COLUMN_WIDTH),
    ],
  });

  return new Table({
    rows: [headerRow, dataRow],
    width: { size: 100, type: WidthType.PERCENTAGE },
    borders: {
      top: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      bottom: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      left: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      right: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      insideHorizontal: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
      insideVertical: { style: BorderStyle.SINGLE, size: 1, color: "000000" },
    },
  });
}

function createReportCell(
  text: string,
  bold: boolean,
  widthPercent: number,
): TableCell {
  return new TableCell({
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text,
            bold,
            size: 24,
            color: "000000",
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 0, before: 0 },
      }),
    ],
    width: { size: widthPercent, type: WidthType.PERCENTAGE },
    margins: {
      top: convertInchesToTwip(0.05),
      right: convertInchesToTwip(0.05),
      bottom: convertInchesToTwip(0.05),
      left: convertInchesToTwip(0.05),
    },
  });
}

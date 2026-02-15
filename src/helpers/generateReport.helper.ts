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

/**
 * Generates a Word document for the patient summary report (from-to date range, table: Jami, Ayollar, Erkaklar, Ishsizlar, Nafaqaxo'rlar, Nogironlar).
 * If from/to are not provided, dateFrom/dateTo from stats (min/max createdAt) are used for the title.
 */
export async function generateReportWord(
  stats: ReportStats,
  dateFrom: Date,
  dateTo: Date,
): Promise<string> {
  const dateFromStr = formatReportDate(dateFrom);
  const dateToStr = formatReportDate(dateTo);
  const title = `${dateFromStr} - ${dateToStr} oraliqda kelgan patsientlar bo'yicha hisobot`;

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

  const now = new Date();
  const dateTime = now
    .toLocaleString("uz-UZ", {
      year: "numeric",
      month: "2-digit",
      day: "2-digit",
      hour: "2-digit",
      minute: "2-digit",
      second: "2-digit",
    })
    .replace(/[/\s:,]+/g, "-");

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

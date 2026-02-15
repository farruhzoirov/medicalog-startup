
import {
  AlignmentType,
  BorderStyle,
  convertInchesToTwip,
  Document,
  Packer,
  PageBreak,
  PageOrientation,
  Paragraph,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType
} from 'docx';
import * as fs from 'fs';
import * as path from 'path';
import * as puppeteer from 'puppeteer';

export const generateWordFile = async (registrations: any[]) => {
  try {
    const doc = new Document({
      sections: [
        {
          properties: {
            page: {
              size: {
                orientation: PageOrientation.LANDSCAPE,
              },
              margin: {
                top: convertInchesToTwip(0.3),
                right: convertInchesToTwip(0.3),
                bottom: convertInchesToTwip(0.3),
                left: convertInchesToTwip(0.3),
              },
            },
          },
          children: [
            // Sahifalarga bo'lingan jadvallar
            ...createPaginatedTables(registrations),
          ],
        },
      ],
    });

    const now = new Date();
    const dateTime = now.toLocaleString('uz-UZ', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    }).replace(/[\/\s:,]+/g, '-');

    const fileName = `medicalog-${dateTime}.docx`;
    const filePath = path.join(process.cwd(), 'uploads', fileName);

    const uploadsDir = path.join(process.cwd(), 'uploads');
    if (!fs.existsSync(uploadsDir)) {
      fs.mkdirSync(uploadsDir, { recursive: true });
    }

    // Generate and save the document
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);

    return `uploads/${fileName}`;
  } catch (error) {
    console.error('Word fayl yaratishda xatolik:', error);
    throw new Error(`Word fayl yaratishda xatolik: ${error.message}`);
  }
};

function createPaginatedTables(registrations: any[]): any[] {
  const itemsPerPage = 35;
  const pages: any[] = [];

  for (let i = 0; i < registrations.length; i += itemsPerPage) {
    const pageRegistrations = registrations.slice(i, i + itemsPerPage);

    pages.push(createRegistrationsTable(pageRegistrations));

    pages.push(
      new Paragraph({
        children: [],
        spacing: { after: 400 },
      })
    );

    if (i + itemsPerPage < registrations.length) {
      pages.push(new PageBreak());
    }
  }

  return pages;
}

function calculateColumnWidths(registrations: any[]): number[] {
  // Header matnlari
  const headers = [
    "Y. t/r",
    "K. t/r",
    "Sana",
    "Ism-familiyasi",
    "T. yili",
    "Fl. r",
    "Kelish sababi",
    "Doza",
    "Manzili",
    "Ish joyi",
    "Xulosa"
  ];

  // Har bir ustun uchun maksimal uzunlikni topish
  const maxLengths = headers.map((header, index) => {
    let maxLength = header.length;

    registrations.forEach(registration => {
      let cellValue = '';

      switch (index) {
        case 0: cellValue = registration.yearlyCount?.toString() || ''; break;
        case 1: cellValue = registration.dailyCount?.toString() || ''; break;
        case 2: cellValue = new Date(registration.createdAt).toLocaleDateString('uz-UZ'); break;
        case 3: cellValue = registration.fullName || ''; break;
        case 4: cellValue = registration.birthDate || ''; break;
        case 5: cellValue = registration.radiologyFilmNumber?.toString() || ''; break;
        case 6: cellValue = registration.visitReason || ''; break;
        case 7: cellValue = registration.radiationDose || ''; break;
        case 8: cellValue = registration.address || ''; break;
        case 9: cellValue = registration.job || ''; break;
        case 10: cellValue = registration.radiologyReport || ''; break;
      }

      maxLength = Math.max(maxLength, cellValue.length);
    });

    return maxLength;
  });

  // Jami uzunlik
  const totalLength = maxLengths.reduce((sum, length) => sum + length, 0);

  // Har bir ustun uchun foiz hisoblash (minimum 5%, maksimum 20%)
  return maxLengths.map(length => {
    const percentage = (length / totalLength) * 100;
    return Math.max(5, Math.min(20, percentage));
  });
}

function createRegistrationsTable(registrations: any[]): Table {
  // Avtomatik width hisoblash
  const columnWidths = calculateColumnWidths(registrations);

  const headerRow = new TableRow({
    children: [
      createHeaderCell("Y", "Y. t/r", columnWidths[0]),
      createHeaderCell("K", "K. t/r", columnWidths[1]),
      createHeaderCell("S", "Sana", columnWidths[2]),
      createHeaderCell("I", "Ism-familiyasi", columnWidths[3]),
      createHeaderCell("T", "T. yili", columnWidths[4]),
      createHeaderCell("F", "Fl. r", columnWidths[5]),
      createHeaderCell("K", "Kelish sababi", columnWidths[6]),
      createHeaderCell("D", "Doza", columnWidths[7]),
      createHeaderCell("M", "Manzili", columnWidths[8]),
      createHeaderCell("I", "Ish joyi", columnWidths[9]),
      createHeaderCell("X", "Xulosa", columnWidths[10]),
    ],
  });

  const dataRows: TableRow[] = registrations.map((registration, index) => {
    const createdDate = new Date(registration.createdAt).toLocaleDateString('uz-UZ');

    return new TableRow({
      children: [
        createDataCell(registration.yearlyCount?.toString() || '', columnWidths[0]),
        createDataCell(registration.dailyCount?.toString() || '', columnWidths[1]),
        createDataCell(createdDate, columnWidths[2]),
        createDataCell(registration.fullName || '', columnWidths[3]),
        createDataCell(registration.birthDate || '', columnWidths[4]),
        createDataCell(registration.radiologyFilmNumber?.toString() || '', columnWidths[5]),
        createDataCell(registration.visitReason || '', columnWidths[6]),
        createDataCell(registration.radiationDose || '', columnWidths[7]),
        createDataCell(registration.address || '', columnWidths[8]),
        createDataCell(registration.job || '', columnWidths[9]),
        createDataCell(registration.radiologyReport || '', columnWidths[10]),
      ],
    });
  });

  return new Table({
    rows: [headerRow, ...dataRows],
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

function createHeaderCell(shortText: string, fullText: string, widthPercent: number): TableCell {
  return new TableCell({
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: fullText,
            bold: true,
            size: 12,
            color: "000000",
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 0, before: 0 },
      }),
    ],
    width: { size: widthPercent, type: WidthType.PERCENTAGE },
    margins: {
      top: convertInchesToTwip(0.02),
      right: convertInchesToTwip(0.02),
      bottom: convertInchesToTwip(0.02),
      left: convertInchesToTwip(0.02),
    },
  });
}

function createDataCell(text: string, widthPercent: number): TableCell {
  return new TableCell({
    children: [
      new Paragraph({
        children: [
          new TextRun({
            text: text || '',
            size: 11,
            color: "000000",
          }),
        ],
        alignment: AlignmentType.CENTER,
        spacing: { after: 0, before: 0 },
      }),
    ],
    width: { size: widthPercent, type: WidthType.PERCENTAGE },
    margins: {
      top: convertInchesToTwip(0.01),
      right: convertInchesToTwip(0.01),
      bottom: convertInchesToTwip(0.01),
      left: convertInchesToTwip(0.01),
    },
  });
}

export const generatePdfFile = async (registrations: any[]) => {
  try {
    const browser = await puppeteer.launch({
      headless: true,
      args: [
        '--no-sandbox',
        '--disable-setuid-sandbox',
        '--disable-dev-shm-usage',
        '--disable-accelerated-2d-canvas',
        '--no-first-run',
        '--no-zygote',
        '--single-process',
        '--disable-gpu'
      ]
    });

    const page = await browser.newPage();

    // HTML content yaratish
    const htmlContent = generateHtmlContent(registrations);

    await page.setContent(htmlContent, { waitUntil: 'networkidle0' });

    // PDF fayl nomi
    const now = new Date();
    const dateTime = now.toLocaleString('uz-UZ', {
      year: 'numeric',
      month: '2-digit',
      day: '2-digit',
      hour: '2-digit',
      minute: '2-digit',
      second: '2-digit'
    }).replace(/[\/\s:,]+/g, '-');

    const fileName = `medicalog-${dateTime}.pdf`;
    const filePath = path.join(process.cwd(), 'uploads', fileName);

    // Uploads papkasini yaratish
    const uploadsDir = path.join(process.cwd(), 'uploads');
    if (!fs.existsSync(uploadsDir)) {
      fs.mkdirSync(uploadsDir, { recursive: true });
    }

    // PDF yaratish
    await page.pdf({
      path: filePath,
      format: 'A4',
      landscape: true,
      margin: {
        top: '0.2in',
        right: '0.2in',
        bottom: '0.2in',
        left: '0.2in'
      },
      printBackground: true
    });

    await browser.close();

    return `uploads/${fileName}`;
  } catch (error) {
    console.error('PDF fayl yaratishda xatolik:', error);
    throw new Error(`PDF fayl yaratishda xatolik: ${error.message}`);
  }
};

function generateHtmlContent(registrations: any[]): string {
  const itemsPerPage = 35;
  let html = `
    <!DOCTYPE html>
    <html>
    <head>
      <meta charset="UTF-8">
      <style>
        body {
          font-family: 'Times New Roman', 'Arial' sans-serif;
          margin: 0;
          padding: 0;
        }
        table {
          width: 100%;
          border-collapse: collapse;
          font-size: 10px;
          border-spacing: 0;
        }
        th, td {
          border: 1px solid #000;
          padding: 4px;
          text-align: center;
          vertical-align: top;
          margin: 0;
        }
        tr {
          border-collapse: collapse;
        }
        th {
          color: #000;
          font-weight: bold;
          font-size: 11px;
        }
        .page-break {
          page-break-before: always;
        }
        @page {
          size: A4 landscape;
          margin: 0.2in;
        }
      </style>
    </head>
    <body>
  `;

  for (let i = 0; i < registrations.length; i += itemsPerPage) {
    const pageRegistrations = registrations.slice(i, i + itemsPerPage);

    if (i > 0) {
      html += '<div class="page-break"></div>';
    }

    html += generateTableHtml(pageRegistrations);
  }

  html += `
    </body>
    </html>
  `;

  return html;
}

function generateTableHtml(registrations: any[]): string {
  let html = `
    <table>
      <thead>
        <tr>
          <th>Y. t/r</th>
          <th>K. t/r</th>
          <th>Sana</th>
          <th>Ism-familiyasi</th>
          <th>T. yili</th>
          <th>Fl. r</th>
          <th>Kelish sababi</th>
          <th>Doza</th>
          <th>Manzili</th>
          <th>Ish joyi</th>
          <th>Xulosa</th>
        </tr>
      </thead>
      <tbody>
  `;

  registrations.forEach(registration => {
    const createdDate = new Date(registration.createdAt).toLocaleDateString('uz-UZ');

    html += `
      <tr>
        <td>${registration.yearlyCount || ''}</td>
        <td>${registration.dailyCount || ''}</td>
        <td>${createdDate}</td>
        <td>${registration.fullName || ''}</td>
        <td>${registration.birthDate || ''}</td>
        <td>${registration.radiologyFilmNumber || ''}</td>
        <td>${registration.visitReason || ''}</td>
        <td>${registration.radiationDose || ''}</td>
        <td>${registration.address || ''}</td>
        <td>${registration.job || ''}</td>
        <td>${registration.radiologyReport || ''}</td>
      </tr>
    `;
  });

  html += `
      </tbody>
    </table>
  `;

  return html;
}


import { NextRequest, NextResponse } from "next/server";
import ExcelJS from "exceljs";

export async function POST(req: NextRequest) {
  try {
    const formData = await req.formData();
    const file = formData.get("file") as File;
    const configStr = formData.get("config") as string;
    const logoTopFile = formData.get("logoTop") as File | null;
    const logoBottomFile = formData.get("logoBottom") as File | null;

    if (!file || !configStr) {
      return NextResponse.json(
        { message: "Missing file or config" },
        { status: 400 }
      );
    }

    const config = JSON.parse(configStr);

    // Read the uploaded Excel file
    const arrayBuffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);

    // Find Data Sheet (First sheet that is not config/output, or just the first one)
    // The user uploads "Employee Data", so likely the first sheet.
    const dataSheet = workbook.worksheets[0];
    if (!dataSheet) {
      return NextResponse.json(
        { message: "No sheets found in uploaded file" },
        { status: 400 }
      );
    }

    // Extract Employees
    const employees = getEmployeeList(dataSheet);
    if (employees.length === 0) {
      return NextResponse.json(
        {
          message:
            "No valid employee data found (Columns: NAME, EPF, DEPARTMENT)",
        },
        { status: 400 }
      );
    }

    // Create Output Workbook
    const outWorkbook = new ExcelJS.Workbook();
    const outSheet = outWorkbook.addWorksheet("Professional Invoices");

    // Setup Columns
    outSheet.getColumn(1).width = 4; // Watermark
    outSheet.getColumn(2).width = 15; // NAME / DEPT Labels
    outSheet.getColumn(3).width = 25; // Description continuation
    outSheet.getColumn(4).width = 8; // Qty
    outSheet.getColumn(5).width = 12; // Price
    outSheet.getColumn(6).width = 15; // Amount

    // Load Images
    let logoTopId: number | undefined;
    let logoBottomId: number | undefined;

    if (logoTopFile) {
      const buffer = await logoTopFile.arrayBuffer();
      logoTopId = outWorkbook.addImage({
        buffer: buffer,
        extension: "png", // Assuming png/jpeg, exceljs handles it
      });
    }
    if (logoBottomFile) {
      const buffer = await logoBottomFile.arrayBuffer();
      logoBottomId = outWorkbook.addImage({
        buffer: buffer,
        extension: "png",
      });
    }

    // Generate Invoices
    let currentEmpIndex = 0;
    let writeRow = 1;
    const invoiceHeight = 25;
    const gap = 5;
    let currentInvoiceNum = Number(config.startInv);

    while (currentEmpIndex < employees.length) {
      // Top Invoice
      const emp1 = employees[currentEmpIndex];
      const invStr1 =
        config.prefix + String(currentInvoiceNum).padStart(5, "0");

      createTemplate(outSheet, writeRow, config);
      fillInvoiceData(outSheet, writeRow, emp1, config, invStr1);

      if (logoTopId !== undefined) {
        outSheet.addImage(logoTopId, {
          tl: { col: 2.5, row: writeRow - 1 + 0.5 }, // Approx position
          ext: { width: 50, height: 50 },
          editAs: "oneCell",
        });
      }
      if (logoBottomId !== undefined) {
        outSheet.addImage(logoBottomId, {
          tl: { col: 1.85, row: writeRow + 19.15 },
          ext: { width: 50, height: 50 },
          editAs: "oneCell",
        });
      }

      currentEmpIndex++;
      currentInvoiceNum++;

      // Bottom Invoice
      if (currentEmpIndex < employees.length) {
        const emp2 = employees[currentEmpIndex];
        const bottomRow = writeRow + invoiceHeight + gap;
        const invStr2 =
          config.prefix + String(currentInvoiceNum).padStart(5, "0");

        createTemplate(outSheet, bottomRow, config);
        fillInvoiceData(outSheet, bottomRow, emp2, config, invStr2);

        if (logoTopId !== undefined) {
          outSheet.addImage(logoTopId, {
            tl: { col: 2.5, row: bottomRow - 1 + 0.5 },
            ext: { width: 50, height: 50 },
            editAs: "oneCell",
          });
        }
        if (logoBottomId !== undefined) {
          outSheet.addImage(logoBottomId, {
            tl: { col: 1.85, row: bottomRow + 19.15 },
            ext: { width: 50, height: 50 },
            editAs: "oneCell",
          });
        }

        // Cut Line
        const cutRow = writeRow + invoiceHeight + 2;
        const cutRowCell = outSheet.getRow(cutRow);
        for (let c = 1; c <= 6; c++) {
          const cell = cutRowCell.getCell(c);
          cell.border = { bottom: { style: "dashed" } };
        }

        currentEmpIndex++;
        currentInvoiceNum++;
      }

      writeRow += 59;
    }

    // Write to buffer
    const buffer = await outWorkbook.xlsx.writeBuffer();

    return new NextResponse(buffer, {
      headers: {
        "Content-Type":
          "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        "Content-Disposition": 'attachment; filename="P&S_main_Invoices.xlsx"',
      },
    });
  } catch (error: any) {
    console.error(error);
    return NextResponse.json({ message: error.message }, { status: 500 });
  }
}

function getEmployeeList(sheet: ExcelJS.Worksheet) {
  const list: any[] = [];
  let headRow = -1,
    nameCol = -1,
    epfCol = -1,
    deptCol = -1;

  // Find Header Row
  sheet.eachRow((row, rowNumber) => {
    if (headRow !== -1) return;

    row.eachCell((cell, colNumber) => {
      const val = String(cell.value || "")
        .toUpperCase()
        .trim();
      if (val.includes("NAME")) {
        headRow = rowNumber;
      }
    });

    if (headRow === rowNumber) {
      row.eachCell((cell, colNumber) => {
        const c = String(cell.value || "")
          .toUpperCase()
          .trim();
        if (c.includes("NAME") && !c.includes("DES")) nameCol = colNumber;
        if (
          c === "EPF" ||
          c === "NO" ||
          c === "ID" ||
          c.includes("EPF") ||
          c.includes("EMP") ||
          c.includes("NUM")
        ) {
          if (epfCol === -1) epfCol = colNumber;
        }
        if (
          c.includes("DEPT") ||
          c.includes("SECTION") ||
          c.includes("LOC") ||
          c.includes("UNIT") ||
          c.includes("BRANCH") ||
          c.includes("DIV") ||
          c.includes("STATION") ||
          c.includes("CATEGORY") ||
          c.includes("COST") ||
          c.includes("CENTER") ||
          c.includes("OFFICE") ||
          c.includes("PLACE") ||
          c.includes("WORK")
        ) {
          deptCol = colNumber;
        }
      });
    }
  });

  if (headRow === -1 || nameCol === -1) return [];

  sheet.eachRow((row, rowNumber) => {
    if (rowNumber <= headRow) return;
    const name = String(row.getCell(nameCol).value || "").trim();
    if (name.length > 1 && name.toUpperCase() !== "NAME") {
      const epf =
        epfCol > -1 ? String(row.getCell(epfCol).value || "").trim() : "";
      const dept =
        deptCol > -1 ? String(row.getCell(deptCol).value || "").trim() : "";
      list.push({ name, epf, dept });
    }
  });

  return list;
}

function createTemplate(sheet: ExcelJS.Worksheet, startRow: number, cfg: any) {
  const r = startRow;
  // Black and White Theme - Pure White Backgrounds
  const THEME_COLOR = "FF000000"; // Black for borders/text
  const TEXT_COLOR = "FF000000"; // Black
  const WHITE = "FFFFFFFF";
  const BORDER_COLOR = "FF000000"; // Black
  const FONT = "Calibri";

  // Set Row Height
  for (let i = 0; i <= 24; i++) {
    sheet.getRow(r + i).height = 15.75;
  }

  // Helper for borders
  const thinBorder: Partial<ExcelJS.Borders> = {
    top: { style: "thin", color: { argb: THEME_COLOR } },
    left: { style: "thin", color: { argb: THEME_COLOR } },
    bottom: { style: "thin", color: { argb: THEME_COLOR } },
    right: { style: "thin", color: { argb: THEME_COLOR } },
  };
  const thickBorder: Partial<ExcelJS.Borders> = {
    top: { style: "medium", color: { argb: THEME_COLOR } },
    left: { style: "medium", color: { argb: THEME_COLOR } },
    bottom: { style: "medium", color: { argb: THEME_COLOR } },
    right: { style: "medium", color: { argb: THEME_COLOR } },
  };

  // Remove internal borders - only outer border will be applied later
  // No border application here

  // Watermark Sidebar (Row r to r+24, Col 1)
  sheet.mergeCells(r, 1, r + 24, 1);
  const sidebar = sheet.getCell(r, 1);
  sidebar.value = (cfg.company || "").toUpperCase() + "  -  OFFICIAL COPY";
  // No fill (White)
  sidebar.alignment = {
    textRotation: 90,
    horizontal: "center",
    vertical: "middle",
  };
  // Text Gray
  sidebar.font = {
    size: 8,
    color: { argb: "FF888888" },
    bold: true,
    name: FONT,
  };
  // Sidebar Border (left edge of invoice)
  sidebar.border = {
    top: { style: "medium", color: { argb: THEME_COLOR } },
    left: { style: "medium", color: { argb: THEME_COLOR } },
    bottom: { style: "medium", color: { argb: THEME_COLOR } },
  };

  // Company Header
  sheet.mergeCells(r + 2, 2, r + 2, 6);
  const compName = sheet.getCell(r + 2, 2);
  compName.value = (cfg.company || "").toUpperCase();
  compName.font = {
    bold: true,
    size: 14,
    color: { argb: THEME_COLOR },
    name: FONT,
  };
  compName.alignment = { horizontal: "center", vertical: "bottom" };

  sheet.mergeCells(r + 3, 2, r + 3, 6);
  const compAddr = sheet.getCell(r + 3, 2);
  compAddr.value = cfg.address;
  compAddr.font = { size: 9, color: { argb: TEXT_COLOR }, name: FONT };
  compAddr.alignment = { horizontal: "center", vertical: "top" };

  // Title
  sheet.mergeCells(r + 4, 2, r + 4, 6);
  const title = sheet.getCell(r + 4, 2);
  title.value = cfg.title;
  title.font = {
    bold: true,
    size: 11,
    color: { argb: TEXT_COLOR },
    name: FONT,
  };
  title.alignment = { horizontal: "center", vertical: "middle" };
  // Bottom border for title section divider
  title.border = {
    bottom: { style: "medium", color: { argb: THEME_COLOR } },
  };

  // Info Row
  sheet.mergeCells(r + 5, 2, r + 5, 6);
  const info = sheet.getCell(r + 5, 2);
  // Value set in fillInvoiceData
  info.font = { size: 10, bold: true, color: { argb: TEXT_COLOR }, name: FONT };
  info.alignment = { horizontal: "center", vertical: "middle" };
  // No fill

  // Greeting & Note
  sheet.mergeCells(r + 7, 2, r + 7, 6);
  const greeting = sheet.getCell(r + 7, 2);
  greeting.value = cfg.greeting;
  greeting.font = {
    bold: true,
    size: 11,
    color: { argb: TEXT_COLOR },
    name: FONT,
  };
  greeting.alignment = { horizontal: "center", vertical: "middle" };

  // Details
  const labelFont = {
    size: 9,
    bold: true,
    color: { argb: THEME_COLOR },
    name: FONT,
  };

  // Name
  sheet.getCell(r + 9, 2).value = "NAME";
  sheet.getCell(r + 9, 2).font = labelFont;
  sheet.mergeCells(r + 9, 3, r + 9, 6);

  // EPF
  sheet.getCell(r + 10, 2).value = "EPF NO";
  sheet.getCell(r + 10, 2).font = labelFont;
  sheet.mergeCells(r + 10, 3, r + 10, 6);

  // Dept
  sheet.getCell(r + 11, 2).value = "DEPARTMENT";
  sheet.getCell(r + 11, 2).font = labelFont;
  sheet.mergeCells(r + 11, 3, r + 11, 6);

  // Table Header
  const headers = ["DESCRIPTION", "QTY", "PRICE", "TOTAL"];
  const headerRow = sheet.getRow(r + 13);

  // Description (Merged Col 2-3)
  sheet.mergeCells(r + 13, 2, r + 13, 3);
  const descHeader = headerRow.getCell(2);
  descHeader.value = headers[0];
  descHeader.font = {
    bold: true,
    size: 9,
    color: { argb: TEXT_COLOR },
    name: FONT,
  };
  descHeader.alignment = { horizontal: "center", vertical: "middle" };
  descHeader.border = {
    top: { style: "medium", color: { argb: THEME_COLOR } },
    left: { style: "medium", color: { argb: THEME_COLOR } },
    bottom: { style: "thin", color: { argb: THEME_COLOR } },
    right: { style: "thin", color: { argb: THEME_COLOR } },
  };
  // Apply border to merged cell 3 as well for consistency if needed, but ExcelJS handles merged borders usually via top-left
  // But we need to ensure the right border of Col 3 is thin
  headerRow.getCell(3).border = {
    top: { style: "medium", color: { argb: THEME_COLOR } },
    left: { style: "thin", color: { argb: THEME_COLOR } },
    bottom: { style: "thin", color: { argb: THEME_COLOR } },
    right: { style: "thin", color: { argb: THEME_COLOR } },
  };

  // QTY (Col 4)
  const qtyHeader = headerRow.getCell(4);
  qtyHeader.value = headers[1];
  qtyHeader.font = {
    bold: true,
    size: 9,
    color: { argb: TEXT_COLOR },
    name: FONT,
  };
  qtyHeader.alignment = { horizontal: "center", vertical: "middle" };
  qtyHeader.border = {
    top: { style: "medium", color: { argb: THEME_COLOR } },
    left: { style: "thin", color: { argb: THEME_COLOR } },
    bottom: { style: "thin", color: { argb: THEME_COLOR } },
    right: { style: "thin", color: { argb: THEME_COLOR } },
  };

  // PRICE (Col 5)
  const priceHeader = headerRow.getCell(5);
  priceHeader.value = headers[2];
  priceHeader.font = {
    bold: true,
    size: 9,
    color: { argb: TEXT_COLOR },
    name: FONT,
  };
  priceHeader.alignment = { horizontal: "center", vertical: "middle" };
  priceHeader.border = {
    top: { style: "medium", color: { argb: THEME_COLOR } },
    left: { style: "thin", color: { argb: THEME_COLOR } },
    bottom: { style: "thin", color: { argb: THEME_COLOR } },
    right: { style: "thin", color: { argb: THEME_COLOR } },
  };

  // TOTAL (Col 6)
  const totalHeader = headerRow.getCell(6);
  totalHeader.value = headers[3];
  totalHeader.font = {
    bold: true,
    size: 9,
    color: { argb: TEXT_COLOR },
    name: FONT,
  };
  totalHeader.alignment = { horizontal: "center", vertical: "middle" };
  totalHeader.border = {
    top: { style: "medium", color: { argb: THEME_COLOR } },
    left: { style: "thin", color: { argb: THEME_COLOR } },
    bottom: { style: "thin", color: { argb: THEME_COLOR } },
    right: { style: "medium", color: { argb: THEME_COLOR } },
  };

  // Table Body (2 rows)
  for (let i = 0; i < 2; i++) {
    const rowNum = r + 14 + i;
    const row = sheet.getRow(rowNum);

    // Merge Description (Col 2-3)
    sheet.mergeCells(rowNum, 2, rowNum, 3);
    const descCell = row.getCell(2);
    descCell.font = { size: 10, color: { argb: TEXT_COLOR }, name: FONT };
    descCell.alignment = { vertical: "middle" };
    descCell.border = {
      top: { style: "thin", color: { argb: THEME_COLOR } },
      left: { style: "medium", color: { argb: THEME_COLOR } },
      bottom: { style: "thin", color: { argb: THEME_COLOR } },
      right: { style: "thin", color: { argb: THEME_COLOR } },
    };
    row.getCell(3).border = {
      top: { style: "thin", color: { argb: THEME_COLOR } },
      left: { style: "thin", color: { argb: THEME_COLOR } },
      bottom: { style: "thin", color: { argb: THEME_COLOR } },
      right: { style: "thin", color: { argb: THEME_COLOR } },
    };

    // QTY (Col 4)
    const qtyCell = row.getCell(4);
    qtyCell.font = { size: 10, color: { argb: TEXT_COLOR }, name: FONT };
    qtyCell.alignment = { vertical: "middle", horizontal: "center" };
    qtyCell.border = {
      top: { style: "thin", color: { argb: THEME_COLOR } },
      left: { style: "thin", color: { argb: THEME_COLOR } },
      bottom: { style: "thin", color: { argb: THEME_COLOR } },
      right: { style: "thin", color: { argb: THEME_COLOR } },
    };

    // PRICE (Col 5)
    const priceCell = row.getCell(5);
    priceCell.font = { size: 10, color: { argb: TEXT_COLOR }, name: FONT };
    priceCell.alignment = { vertical: "middle", horizontal: "right" };
    priceCell.border = {
      top: { style: "thin", color: { argb: THEME_COLOR } },
      left: { style: "thin", color: { argb: THEME_COLOR } },
      bottom: { style: "thin", color: { argb: THEME_COLOR } },
      right: { style: "thin", color: { argb: THEME_COLOR } },
    };

    // TOTAL (Col 6)
    const totalCell = row.getCell(6);
    totalCell.font = { size: 10, color: { argb: TEXT_COLOR }, name: FONT };
    totalCell.alignment = { vertical: "middle", horizontal: "right" };
    totalCell.border = {
      top: { style: "thin", color: { argb: THEME_COLOR } },
      left: { style: "thin", color: { argb: THEME_COLOR } },
      bottom: { style: "thin", color: { argb: THEME_COLOR } },
      right: { style: "medium", color: { argb: THEME_COLOR } },
    };
  }

  // Add medium bottom border to last table row (r+15)
  const lastRow = sheet.getRow(r + 15);
  lastRow.getCell(2).border = {
    ...lastRow.getCell(2).border,
    bottom: { style: "medium", color: { argb: THEME_COLOR } },
    left: { style: "medium", color: { argb: THEME_COLOR } },
  };
  lastRow.getCell(3).border = {
    ...lastRow.getCell(3).border,
    bottom: { style: "medium", color: { argb: THEME_COLOR } },
  };
  lastRow.getCell(4).border = {
    ...lastRow.getCell(4).border,
    bottom: { style: "medium", color: { argb: THEME_COLOR } },
  };
  lastRow.getCell(5).border = {
    ...lastRow.getCell(5).border,
    bottom: { style: "medium", color: { argb: THEME_COLOR } },
  };
  lastRow.getCell(6).border = {
    ...lastRow.getCell(6).border,
    bottom: { style: "medium", color: { argb: THEME_COLOR } },
    right: { style: "medium", color: { argb: THEME_COLOR } },
  };

  // Empty row after table - no borders needed

  // Grand Total
  sheet.mergeCells(r + 17, 4, r + 18, 5); // Merge QTY, PRICE for Label
  const gtLabel = sheet.getCell(r + 17, 4);
  gtLabel.value = "GRAND TOTAL";
  gtLabel.font = {
    bold: true,
    size: 9,
    color: { argb: THEME_COLOR },
    name: FONT,
  };
  gtLabel.alignment = { horizontal: "center", vertical: "middle" };
  // No border for Grand Total label

  sheet.mergeCells(r + 17, 6, r + 18, 6); // TOTAL Value
  const gtVal = sheet.getCell(r + 17, 6);
  // Text Black
  gtVal.font = {
    bold: true,
    size: 12,
    color: { argb: TEXT_COLOR },
    name: FONT,
  };
  gtVal.alignment = { horizontal: "center", vertical: "middle" };
  // Border around Grand Total value
  gtVal.border = {
    top: { style: "thin", color: { argb: THEME_COLOR } },
    left: { style: "thin", color: { argb: THEME_COLOR } },
    bottom: { style: "thin", color: { argb: THEME_COLOR } },
    right: { style: "thin", color: { argb: THEME_COLOR } },
  };

  // Footer
  sheet.mergeCells(r + 19, 2, r + 19, 6);
  const footer = sheet.getCell(r + 19, 2);
  footer.value = "VALID UNTIL: " + cfg.valid;
  footer.alignment = { horizontal: "center" };
  footer.font = {
    bold: true,
    size: 9,
    color: { argb: TEXT_COLOR },
    name: FONT,
  };

  // Logo Box (Merged)
  sheet.mergeCells(r + 20, 2, r + 22, 2);

  // Signature
  sheet.mergeCells(r + 23, 4, r + 23, 6);
  const sig = sheet.getCell(r + 23, 4);
  sig.value = "_________________________\nAUTHORIZED SIGNATURE";
  sig.alignment = { horizontal: "right", vertical: "bottom", wrapText: true };
  sig.font = { size: 8, color: { argb: TEXT_COLOR }, name: FONT };

  // Terms
  sheet.mergeCells(r + 24, 2, r + 24, 6);
  const terms = sheet.getCell(r + 24, 2);
  terms.value = cfg.terms;
  terms.font = {
    size: 12,
    bold: true,
    italic: true,
    color: { argb: "FF000000" },
    name: FONT,
  };
  terms.alignment = {
    horizontal: "center",
    vertical: "middle",
    wrapText: true,
  };

  // Outer Border (Outline only)
  // Top
  for (let c = 2; c <= 6; c++)
    sheet.getCell(r, c).border = {
      ...sheet.getCell(r, c).border,
      top: { style: "medium", color: { argb: THEME_COLOR } },
    };
  // Bottom
  for (let c = 2; c <= 6; c++)
    sheet.getCell(r + 24, c).border = {
      ...sheet.getCell(r + 24, c).border,
      bottom: { style: "medium", color: { argb: THEME_COLOR } },
    };
  // Right edge (column 6)
  for (let row = r; row <= r + 24; row++)
    sheet.getCell(row, 6).border = {
      ...sheet.getCell(row, 6).border,
      right: { style: "medium", color: { argb: THEME_COLOR } },
    };
}

function fillInvoiceData(
  sheet: ExcelJS.Worksheet,
  r: number,
  emp: any,
  cfg: any,
  invStr: string
) {
  const FONT = "Calibri";
  const TEXT_COLOR = "FF333333";

  sheet.getCell(r + 5, 2).value =
    "INVOICE NO: " + invStr + "   |   DATE: " + cfg.date;

  const detailFont = { size: 12, color: { argb: TEXT_COLOR }, name: FONT };

  sheet.getCell(r + 9, 3).value = ": " + emp.name;
  sheet.getCell(r + 9, 3).font = detailFont;

  sheet.getCell(r + 10, 3).value = ": " + emp.epf;
  sheet.getCell(r + 10, 3).font = detailFont;

  let d = emp.dept;
  if (!d || d === "" || d === "N/A") d = "-";
  sheet.getCell(r + 11, 3).value = ": " + d;
  sheet.getCell(r + 11, 3).font = detailFont;

  const fmt = (n: any) =>
    "Rs " +
    Number(n).toLocaleString("en-US", {
      minimumFractionDigits: 2,
      maximumFractionDigits: 2,
    });

  const total1 = Number(cfg.price1) * Number(cfg.qty1);
  const total2 = Number(cfg.price2) * Number(cfg.qty2);
  const grandTotal = total1 + total2;

  if (cfg.item1) {
    // sheet.getCell(r + 14, 2).value = "01"; // NO - Removed
    // sheet.mergeCells(r + 14, 2, r + 14, 3); // Already merged in createTemplate
    sheet.getCell(r + 14, 2).value = cfg.item1; // DESC
    sheet.getCell(r + 14, 4).value = cfg.qty1; // QTY
    sheet.getCell(r + 14, 5).value = fmt(cfg.price1); // PRICE
    sheet.getCell(r + 14, 6).value = fmt(total1); // AMOUNT

    sheet.getCell(r + 14, 2).alignment = {
      horizontal: "left",
      vertical: "middle",
    }; // Left align desc
    sheet.getCell(r + 14, 4).alignment = {
      horizontal: "center",
      vertical: "middle",
    };
    sheet.getCell(r + 14, 6).alignment = {
      horizontal: "right",
      vertical: "middle",
    };
  }

  if (cfg.item2) {
    // sheet.getCell(r + 15, 2).value = "02"; // NO - Removed
    // sheet.mergeCells(r + 15, 2, r + 15, 3); // Already merged in createTemplate
    sheet.getCell(r + 15, 2).value = cfg.item2; // DESC
    sheet.getCell(r + 15, 4).value = cfg.qty2; // QTY
    sheet.getCell(r + 15, 5).value = fmt(cfg.price2); // PRICE
    sheet.getCell(r + 15, 6).value = fmt(total2); // AMOUNT

    sheet.getCell(r + 15, 2).alignment = {
      horizontal: "left",
      vertical: "middle",
    };
    sheet.getCell(r + 15, 4).alignment = {
      horizontal: "center",
      vertical: "middle",
    };
    sheet.getCell(r + 15, 6).alignment = {
      horizontal: "right",
      vertical: "middle",
    };
  }

  sheet.getCell(r + 17, 6).value = fmt(grandTotal);
}

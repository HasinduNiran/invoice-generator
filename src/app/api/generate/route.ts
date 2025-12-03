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
    outSheet.getColumn(1).width = 4; // Watermark (approx 25px)
    outSheet.getColumn(2).width = 50; // Description (approx 350px)
    outSheet.getColumn(3).width = 6; // Qty (approx 40px)
    outSheet.getColumn(4).width = 13; // Price (approx 90px)
    outSheet.getColumn(5).width = 13; // Total (approx 90px)

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
    const invoiceHeight = 29;
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
        });
      }
      if (logoBottomId !== undefined) {
        outSheet.addImage(logoBottomId, {
          tl: { col: 1.2, row: writeRow + 20 }, // Approx position
          ext: { width: 50, height: 50 },
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
          });
        }
        if (logoBottomId !== undefined) {
          outSheet.addImage(logoBottomId, {
            tl: { col: 1.2, row: bottomRow + 20 },
            ext: { width: 50, height: 50 },
          });
        }

        // Cut Line
        const cutRow = writeRow + invoiceHeight + 2;
        const cutRowCell = outSheet.getRow(cutRow);
        for (let c = 1; c <= 5; c++) {
          const cell = cutRowCell.getCell(c);
          cell.border = { bottom: { style: "dashed" } };
        }

        currentEmpIndex++;
        currentInvoiceNum++;
      }

      writeRow += 67;
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
  for (let i = 0; i <= 28; i++) {
    sheet.getRow(r + i).height = 15.75;
  }

  // Helper for borders
  const borderStyle: Partial<ExcelJS.Borders> = {
    top: { style: "medium", color: { argb: THEME_COLOR } },
    left: { style: "medium", color: { argb: THEME_COLOR } },
    bottom: { style: "medium", color: { argb: THEME_COLOR } },
    right: { style: "medium", color: { argb: THEME_COLOR } },
  };

  // Watermark Sidebar (Row r to r+28, Col 1)
  sheet.mergeCells(r, 1, r + 28, 1);
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

  // Company Header
  sheet.mergeCells(r + 2, 2, r + 2, 5);
  const compName = sheet.getCell(r + 2, 2);
  compName.value = (cfg.company || "").toUpperCase();
  compName.font = {
    bold: true,
    size: 14,
    color: { argb: THEME_COLOR },
    name: FONT,
  };
  compName.alignment = { horizontal: "center", vertical: "bottom" };

  sheet.mergeCells(r + 3, 2, r + 3, 5);
  const compAddr = sheet.getCell(r + 3, 2);
  compAddr.value = cfg.address;
  compAddr.font = { size: 9, color: { argb: TEXT_COLOR }, name: FONT };
  compAddr.alignment = { horizontal: "center", vertical: "top" };

  // Title
  sheet.mergeCells(r + 4, 2, r + 4, 5);
  const title = sheet.getCell(r + 4, 2);
  title.value = cfg.title;
  title.font = {
    bold: true,
    size: 11,
    color: { argb: TEXT_COLOR },
    name: FONT,
  };
  title.alignment = { horizontal: "center" };
  title.border = { bottom: { style: "thick", color: { argb: THEME_COLOR } } };

  // Info Row
  sheet.mergeCells(r + 5, 2, r + 5, 5);
  const info = sheet.getCell(r + 5, 2);
  // Value set in fillInvoiceData
  info.font = { size: 10, bold: true, color: { argb: TEXT_COLOR }, name: FONT };
  info.alignment = { horizontal: "center", vertical: "middle" };
  // No fill

  // Greeting & Note
  sheet.mergeCells(r + 7, 2, r + 7, 5);
  const greeting = sheet.getCell(r + 7, 2);
  greeting.value = cfg.greeting;
  greeting.font = {
    bold: true,
    size: 11,
    color: { argb: TEXT_COLOR },
    name: FONT,
  };
  greeting.alignment = { horizontal: "center" };

  // Details
  const labelFont = {
    size: 9,
    bold: true,
    color: { argb: THEME_COLOR },
    name: FONT,
  };
  const valueBorder: Partial<ExcelJS.Borders> = {
    bottom: { style: "dotted", color: { argb: BORDER_COLOR } },
  };

  sheet.getCell(r + 9, 2).value = "NAME";
  sheet.getCell(r + 9, 2).font = labelFont;
  sheet.mergeCells(r + 9, 3, r + 9, 5);
  sheet.getCell(r + 9, 3).border = valueBorder;

  sheet.getCell(r + 10, 2).value = "EPF NO";
  sheet.getCell(r + 10, 2).font = labelFont;
  sheet.mergeCells(r + 10, 3, r + 10, 5);
  sheet.getCell(r + 10, 3).border = valueBorder;

  sheet.getCell(r + 11, 2).value = "DEPARTMENT";
  sheet.getCell(r + 11, 2).font = labelFont;
  sheet.mergeCells(r + 11, 3, r + 11, 5);
  sheet.getCell(r + 11, 3).border = valueBorder;

  // Table Header
  const headers = ["DESCRIPTION", "QTY", "PRICE", "TOTAL"];
  const headerRow = sheet.getRow(r + 13);
  for (let i = 0; i < 4; i++) {
    const cell = headerRow.getCell(i + 2);
    cell.value = headers[i];
    // Text Black
    cell.font = {
      bold: true,
      size: 9,
      color: { argb: TEXT_COLOR },
      name: FONT,
    };
    cell.alignment = { horizontal: "center", vertical: "middle" };
    // No fill
    cell.border = {
      top: { style: "thin", color: { argb: THEME_COLOR } },
      bottom: { style: "thin", color: { argb: THEME_COLOR } },
      left: { style: "thin", color: { argb: THEME_COLOR } },
      right: { style: "thin", color: { argb: THEME_COLOR } },
    };
  }

  // Table Body (2 rows)
  for (let i = 0; i < 2; i++) {
    const row = sheet.getRow(r + 14 + i);
    for (let j = 0; j < 4; j++) {
      const cell = row.getCell(j + 2);
      cell.font = { size: 10, color: { argb: TEXT_COLOR }, name: FONT };
      cell.alignment = { vertical: "middle" };
      cell.border = {
        bottom: { style: "thin", color: { argb: BORDER_COLOR } },
        left: { style: "thin", color: { argb: BORDER_COLOR } },
        right: { style: "thin", color: { argb: BORDER_COLOR } },
      };
    }
  }

  // Ensure empty row after table also has borders if needed, or just rely on outer border
  // But user said "last line not printed", maybe referring to the bottom of the table section
  // Let's add a bottom border to the row after the items to close the table visually if it's separate
  // Actually, the outer border handles the main box.
  // If the user means the table grid, let's ensure the last item row has a strong bottom border
  // The loop above sets 'thin' bottom border.
  // Let's check if there is a gap before Grand Total.
  // Grand Total is at r+17. Table ends at r+15. r+16 is empty.
  // Let's add borders to r+16 to make it look continuous or close the table.
  const emptyRow = sheet.getRow(r + 16);
  for (let j = 0; j < 4; j++) {
    const cell = emptyRow.getCell(j + 2);
    cell.border = {
      left: { style: "thin", color: { argb: BORDER_COLOR } },
      right: { style: "thin", color: { argb: BORDER_COLOR } },
      bottom: { style: "thin", color: { argb: BORDER_COLOR } },
    };
  }

  // Grand Total
  sheet.mergeCells(r + 17, 4, r + 18, 4);
  const gtLabel = sheet.getCell(r + 17, 4);
  gtLabel.value = "GRAND TOTAL";
  gtLabel.font = {
    bold: true,
    size: 9,
    color: { argb: THEME_COLOR },
    name: FONT,
  };
  gtLabel.alignment = { horizontal: "right", vertical: "middle" };

  sheet.mergeCells(r + 17, 5, r + 18, 5);
  const gtVal = sheet.getCell(r + 17, 5);
  // Text Black
  gtVal.font = {
    bold: true,
    size: 14,
    color: { argb: TEXT_COLOR },
    name: FONT,
  };
  gtVal.alignment = { horizontal: "center", vertical: "middle" };
  // No fill
  gtVal.border = {
    top: { style: "medium", color: { argb: THEME_COLOR } },
    bottom: { style: "medium", color: { argb: THEME_COLOR } },
    left: { style: "medium", color: { argb: THEME_COLOR } },
    right: { style: "medium", color: { argb: THEME_COLOR } },
  };

  // Footer
  sheet.mergeCells(r + 20, 2, r + 20, 5);
  const footer = sheet.getCell(r + 20, 2);
  footer.value = "VALID UNTIL: " + cfg.valid;
  footer.alignment = { horizontal: "center" };
  footer.font = {
    bold: true,
    size: 9,
    color: { argb: TEXT_COLOR },
    name: FONT,
  };

  // Signature
  sheet.mergeCells(r + 24, 2, r + 24, 5);
  const sig = sheet.getCell(r + 24, 2);
  sig.value = "_________________________\nAUTHORIZED SIGNATURE";
  sig.alignment = { horizontal: "right", vertical: "bottom", wrapText: true };
  sig.font = { size: 8, color: { argb: TEXT_COLOR }, name: FONT };

  // Terms
  sheet.mergeCells(r + 28, 2, r + 28, 5);
  const terms = sheet.getCell(r + 28, 2);
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

  // Developer Credit
  // sheet.mergeCells(r + 26, 1, r + 26, 5);
  // const dev = sheet.getCell(r + 26, 1);
  // dev.value = "Developed by Flowiix (pvt) LTD";
  // dev.alignment = { horizontal: "center", vertical: "middle" };
  // dev.font = { size: 7, color: { argb: "FF888888" }, name: FONT };

  // Outer Border (Manual)
  // Top
  for (let c = 1; c <= 5; c++)
    sheet.getCell(r, c).border = {
      ...sheet.getCell(r, c).border,
      top: { style: "medium", color: { argb: THEME_COLOR } },
    };
  // Bottom
  for (let c = 1; c <= 5; c++)
    sheet.getCell(r + 28, c).border = {
      ...sheet.getCell(r + 28, c).border,
      bottom: { style: "medium", color: { argb: THEME_COLOR } },
    };
  // Left
  for (let row = r; row <= r + 28; row++)
    sheet.getCell(row, 1).border = {
      ...sheet.getCell(row, 1).border,
      left: { style: "medium", color: { argb: THEME_COLOR } },
    };
  // Right
  for (let row = r; row <= r + 28; row++)
    sheet.getCell(row, 5).border = {
      ...sheet.getCell(row, 5).border,
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
    sheet.getCell(r + 14, 2).value = cfg.item1;
    sheet.getCell(r + 14, 3).value = cfg.qty1;
    sheet.getCell(r + 14, 4).value = fmt(cfg.price1);
    sheet.getCell(r + 14, 5).value = fmt(total1);

    sheet.getCell(r + 14, 3).alignment = {
      horizontal: "center",
      vertical: "middle",
    };
    sheet.getCell(r + 14, 5).alignment = {
      horizontal: "right",
      vertical: "middle",
    };
  }

  if (cfg.item2) {
    sheet.getCell(r + 15, 2).value = cfg.item2;
    sheet.getCell(r + 15, 3).value = cfg.qty2;
    sheet.getCell(r + 15, 4).value = fmt(cfg.price2);
    sheet.getCell(r + 15, 5).value = fmt(total2);

    sheet.getCell(r + 15, 3).alignment = {
      horizontal: "center",
      vertical: "middle",
    };
    sheet.getCell(r + 15, 5).alignment = {
      horizontal: "right",
      vertical: "middle",
    };
  }

  sheet.getCell(r + 17, 5).value = fmt(grandTotal);
}

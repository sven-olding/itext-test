package com.gi.itext;

import com.itextpdf.kernel.geom.PageSize;
import com.itextpdf.kernel.pdf.PdfDocument;
import com.itextpdf.kernel.pdf.PdfWriter;
import com.itextpdf.layout.Document;
import com.itextpdf.layout.borders.SolidBorder;
import com.itextpdf.layout.element.AreaBreak;
import com.itextpdf.layout.element.Paragraph;
import com.itextpdf.layout.element.Table;
import com.itextpdf.layout.property.AreaBreakType;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

public class App {
    private static final String OUTPUT = "./target/output.pdf";
    private static final String INPUT = "/input.xlsx";
    private static final float PT_TO_CM = 0.0352778f;

    public static void main(String[] args) {
        new App().run();
    }

    public void run() {
        File file = new File(OUTPUT);
        file.getParentFile().mkdirs();

        try (InputStream is = getClass().getResourceAsStream(INPUT);
             Workbook workbook = new XSSFWorkbook(is)) {
            PdfDocument pdfDoc;
            try {
                pdfDoc = new PdfDocument(new PdfWriter(OUTPUT));
            } catch (FileNotFoundException e) {
                e.printStackTrace();
                return;
            }
            pdfDoc.addNewPage();
            Document doc = new Document(pdfDoc, PageSize.A4);

            for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
                Sheet sheet = workbook.getSheetAt(sheetNum);
                if (sheetNum > 0) {
                    doc.add(new AreaBreak(AreaBreakType.NEXT_PAGE));
                }
                int numCols = getNumberOfColumns(sheet);
                float[] columnWidth = new float[numCols];
                for (int j = 0; j < numCols; j++) {
                    float columnWidthInPixels = sheet.getColumnWidthInPixels(j);
                    double columnWidthInPoints = columnWidthInPixels * 0.75d;
                    columnWidth[j] = (float) columnWidthInPoints;
                }
                Table table = new Table(columnWidth);
                table.useAllAvailableWidth();

                for (int rowNum = 0; rowNum <= sheet.getLastRowNum(); rowNum++) {
                    Row row = sheet.getRow(rowNum);
                    if (row == null) {
                        continue;
                    }
                    float heightInPoints = row.getHeightInPoints();
                    System.out.println("Row: " + (rowNum + 1) + ": " + heightInPoints + "pt = " + heightInPoints * PT_TO_CM + "cm");
                    for (int cellNum = 0; cellNum < numCols; cellNum++) {
                        Cell cell = row.getCell(cellNum, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK);
                        com.itextpdf.layout.element.Cell pdfCell = new com.itextpdf.layout.element.Cell();
                        XSSFCellStyle cellStyle = (XSSFCellStyle) cell.getCellStyle();
                        pdfCell.add(
                                new Paragraph(cell.getStringCellValue())
                                        .setFontSize(cellStyle.getFont().getFontHeightInPoints()));
                        pdfCell.setHeight(heightInPoints);
                        pdfCell.setBorder(new SolidBorder(0.5f));
                        table.addCell(pdfCell);
                    }
                }
                doc.add(table);
            }

            doc.close();
            pdfDoc.close();
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    private int getNumberOfColumns(Sheet sheet) {
        int firstRowNum = sheet.getFirstRowNum();
        int lastRowNum = sheet.getLastRowNum();
        for (int rowNum = firstRowNum; rowNum < lastRowNum; rowNum++) {
            Row row = sheet.getRow(rowNum);
            if (row == null) {
                continue;
            }
            if (row.getLastCellNum() > -1) {
                return row.getLastCellNum();
            }
        }
        return -1;
    }
}

/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.kles.output.xls;

import com.kles.output.AbstractOutput;
import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Locale;
import java.util.Map;
import java.util.ResourceBundle;
import org.apache.poi.hssf.usermodel.HSSFPalette;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author jchau
 */
public abstract class ExcelOutput extends AbstractOutput {

    public static int XLS_TABLE_HORIZONTAL = 0;
    public static int XLS_TABLE_VERTICAL = 1;
    public static int XLS_FORMAT = 0;
    public static int XLSX_FORMAT = 1;

    public static final String HEADER_STYLE = "header";
    public static final String LEFT_BOLD_STYLE = "cell_bold_left";
    public static final String CENTER_STYLE = "cell_centered";
    public static final String CENTER_LOCK_STYLE = "cell_centered_locked";
    public static final String CENTER_BOLD_STYLE = "cell_bold_centered";
    public static final String CENTER_BOLD_LOCK_STYLE = "cell_bold_centered_locked";
    protected Locale local = Locale.getDefault();
    protected ResourceBundle resourceMessage;
    protected Workbook workbook;
    protected Map<String, CellStyle> styles;

    public ExcelOutput() {
        this(0);
    }

    public ExcelOutput(int type) {
        if (type == XLS_FORMAT) {
            workbook = new HSSFWorkbook();
        } else {
            workbook = new XSSFWorkbook();
        }
        styles = createStyles(workbook);
    }

    @Override
    public boolean save() {
        try {
            if (workbook instanceof HSSFWorkbook) {
                if (!filepath.endsWith("xls")) {
                    filepath += ".xls";
                }
            } else if (workbook instanceof XSSFWorkbook) {
                if (!filepath.endsWith("xlsx")) {
                    filepath += ".xlsx";
                }
            }
            try (FileOutputStream out = new FileOutputStream(new File(filepath))) {
                workbook.write(out);
                out.close();
            }
        } catch (IOException e) {
            return false;
        }
        return true;
    }

    public void recalculate() {
        FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
        for (int sheetNum = 0; sheetNum < workbook.getNumberOfSheets(); sheetNum++) {
            Sheet sheet = workbook.getSheetAt(sheetNum);
            for (Row r : sheet) {
                for (Cell c : r) {
                    if (c.getCellType() == Cell.CELL_TYPE_FORMULA) {
                        evaluator.evaluateFormulaCell(c);
                    }
                    sheet.autoSizeColumn(c.getColumnIndex());
                }
            }
        }

    }

    protected Cell addCell(Row row, int current) {
        return row.createCell(current++);
    }

    public Sheet writeXlsTable(Sheet sheet, String title, String[] header, String[][] data, int beginROW, int beginCOL) {
        sheet.setDisplayGridlines(false);
        sheet.setPrintGridlines(false);
        sheet.setFitToPage(true);
        sheet.setHorizontallyCenter(true);
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);

        int rownum = beginROW;
        int cellnum = beginCOL;
        Row row = sheet.createRow(rownum++);
//        Cell cellTitle = row.createCell(cellnum);
//        cellTitle.setCellValue(title);
//        cellTitle.setCellStyle(styles.get("cell_bold_centered_locked"));
//        sheet.addMergedRegion(new CellRangeAddress(rownum, rownum, cellnum, cellnum + header.length));
//
//        row = sheet.createRow(rownum++);
        for (int k = 0; k < header.length; k++) {
            Cell cell = row.createCell(cellnum++);
            cell.setCellValue(header[k]);
            cell.setCellStyle(styles.get("header"));
            //rownum++;
//            row = sheet.createRow(rownum);
        }
        for (String[] tab : data) {
            cellnum = beginCOL;
            row = sheet.createRow(rownum++);
            for (int k = 0; k < tab.length; k++) {
                Cell cell = row.createCell(cellnum++);
                cell.setCellValue(tab[k]);
                cell.setCellStyle(styles.get("cell_normal_centered"));
                sheet.autoSizeColumn(k);
            }
        }
        return sheet;
    }

    /**
     * create a library of cell styles
     *
     * @param wb
     * @return
     */
    public static Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<>();
        DataFormat df = wb.createDataFormat();

        Font font1 = wb.createFont();

        HSSFPalette palette = ((HSSFWorkbook) wb).getCustomPalette();
        short colorIndex = 45;
        palette.setColorAtIndex(colorIndex, (byte) 54, (byte) 96, (byte) 146);

        CellStyle style;
        Font headerFont = wb.createFont();
        headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        headerFont.setColor(IndexedColors.WHITE.getIndex());

        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFillForegroundColor((short) 45);
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFont(headerFont);
        styles.put(HEADER_STYLE, style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFont(font1);
        style.setLocked(false);
        styles.put(CENTER_STYLE, style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFont(font1);
        style.setLocked(true);
        styles.put(CENTER_LOCK_STYLE, style);

        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFillForegroundColor(IndexedColors.BLUE_GREY.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFont(headerFont);
        style.setDataFormat(df.getFormat("d-mmm"));
        styles.put("header_date", style);

        font1.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_LEFT);
        style.setFont(font1);
        styles.put(LEFT_BOLD_STYLE, style);

        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFont(font1);
        style.setLocked(false);
        styles.put(CENTER_BOLD_STYLE, style);

        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFont(font1);
        style.setLocked(true);
        styles.put(CENTER_BOLD_LOCK_STYLE, style);

        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_RIGHT);
        style.setFont(font1);
        style.setDataFormat(df.getFormat("dd/mm/yyyy hh:mm:ss"));
        styles.put("cell_b_date", style);

        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_RIGHT);
        style.setFont(font1);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setDataFormat(df.getFormat("d-mmm"));
        styles.put("cell_g", style);

        Font font2 = wb.createFont();
        font2.setColor(IndexedColors.BLUE.getIndex());
        font2.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_LEFT);
        style.setFont(font2);
        styles.put("cell_bb", style);

        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_RIGHT);
        style.setFont(font1);
        style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setDataFormat(df.getFormat("d-mmm"));
        styles.put("cell_bg", style);

        Font font3 = wb.createFont();
        font3.setFontHeightInPoints((short) 14);
        font3.setColor(IndexedColors.DARK_BLUE.getIndex());
        font3.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_LEFT);
        style.setFont(font3);
        style.setWrapText(true);
        styles.put("cell_h", style);

        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_LEFT);
        style.setWrapText(true);
        styles.put("cell_normal", style);

        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setWrapText(true);
        styles.put("cell_normal_centered", style);

        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_RIGHT);
        style.setWrapText(true);
        style.setDataFormat(df.getFormat("d-mmm"));
        styles.put("cell_normal_date", style);

        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_LEFT);
        style.setIndention((short) 1);
        style.setWrapText(true);
        styles.put("cell_indented", style);

        style = createBorderedStyle(wb);
        style.setFillForegroundColor(IndexedColors.BLUE.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        styles.put("cell_blue", style);

        return styles;
    }

    public static CellStyle createBorderedStyle(Workbook wb) {
        CellStyle style = wb.createCellStyle();
        style.setBorderRight(CellStyle.BORDER_THIN);
        style.setRightBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderBottom(CellStyle.BORDER_THIN);
        style.setBottomBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderLeft(CellStyle.BORDER_THIN);
        style.setLeftBorderColor(IndexedColors.BLACK.getIndex());
        style.setBorderTop(CellStyle.BORDER_THIN);
        style.setTopBorderColor(IndexedColors.BLACK.getIndex());
        return style;
    }

    public ResourceBundle getResourceMessage() {
        return resourceMessage;
    }

    public void setResourceMessage(ResourceBundle resourceMessage) {
        this.resourceMessage = resourceMessage;
    }
}

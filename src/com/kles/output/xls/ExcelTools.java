package com.kles.output.xls;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.text.DateFormat;
import java.util.Date;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;
import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.PrintSetup;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * **********************************************************************
 Définition d'un ExcelTools en écriture
 *
 * @author Jérémy Chaut
 * *********************************************************************
 */
public class ExcelTools {

    private static DateFormat dateFormat;
    private static Date date;
    public static int XLS_TABLE_HORIZONTAL = 0;
    public static int XLS_TABLE_VERTICAL = 1;
    private static Map<String, CellStyle> styles;
    private Workbook workbook;

    /**
     * *********************************************************************
     * Constructeur de la classe Excel
     * *********************************************************************
     * @param wb
     */
    public ExcelTools(Workbook wb) {
        styles = createStyles(wb);
    }

    public ExcelTools() {
        InputStream in = null;
        try {
            FileUtils.copyURLToFile(getClass().getClassLoader().getResource("template" + System.getProperty("file.separator") + "m3.xls"), new File("temp.xls"));
            in = new java.io.FileInputStream(new File("temp.xls"));
            workbook = new HSSFWorkbook(in);
            styles = createStyles(workbook);
        } catch (FileNotFoundException ex) {
            Logger.getLogger(ExcelTools.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IOException ex) {
            Logger.getLogger(ExcelTools.class.getName()).log(Level.SEVERE, null, ex);
        } finally {
            try {
                in.close();
            } catch (IOException ex) {
                Logger.getLogger(ExcelTools.class.getName()).log(Level.SEVERE, null, ex);
            }
        }
    }

    /**
     * create a library of cell styles
     */
    private static Map<String, CellStyle> createStyles(Workbook wb) {
        Map<String, CellStyle> styles = new HashMap<>();
        DataFormat df = wb.createDataFormat();

        Font font1 = wb.createFont();

        CellStyle style;
        Font headerFont = wb.createFont();
        headerFont.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
        style.setFont(headerFont);
        styles.put("header", style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFont(font1);
        style.setLocked(false);
        styles.put("cell_centered", style);

        style = wb.createCellStyle();
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFont(font1);
        style.setLocked(true);
        styles.put("cell_centered_locked", style);
//        style = createBorderedStyle(wb);
//        style.setAlignment(CellStyle.ALIGN_CENTER);
//        style.setFillForegroundColor(IndexedColors.LIGHT_CORNFLOWER_BLUE.getIndex());
//        style.setFillPattern(CellStyle.SOLID_FOREGROUND);
//        style.setFont(headerFont);
//        style.setDataFormat(df.getFormat("d-mmm"));
//        styles.put("header_date", style);
        font1.setBoldweight(Font.BOLDWEIGHT_BOLD);
        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_LEFT);
        style.setFont(font1);
        styles.put("cell_b", style);

        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFont(font1);
        style.setLocked(false);
        styles.put("cell_b_centered", style);

        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_CENTER);
        style.setFont(font1);
        style.setLocked(true);
        styles.put("cell_b_centered_locked", style);

        style = createBorderedStyle(wb);
        style.setAlignment(CellStyle.ALIGN_RIGHT);
        style.setFont(font1);
        style.setDataFormat(df.getFormat("d-mmm"));
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

    private static CellStyle createBorderedStyle(Workbook wb) {
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

    public static Sheet writeXlsTable(Sheet sheet, String title, String[] header, String[][] data, int beginROW, int beginCOL) {
        sheet.setDisplayGridlines(false);
        sheet.setPrintGridlines(false);
        sheet.setFitToPage(true);
        sheet.setHorizontallyCenter(true);
        PrintSetup printSetup = sheet.getPrintSetup();
        printSetup.setLandscape(true);

        int rownum = beginROW;
        int cellnum = beginCOL;
        Row rowTitle = sheet.createRow(rownum);
        Cell cellTitle = rowTitle.createCell(cellnum);
        cellTitle.setCellValue(title);
        cellTitle.setCellStyle(styles.get("cell_b_centered_locked"));
        sheet.addMergedRegion(new CellRangeAddress(rownum, rownum, cellnum, cellnum + 1));

        rownum++;
        Row row = sheet.createRow(rownum);
        for (int k = 0; k < header.length; k++) {
            Cell cell = row.createCell(cellnum);
            cell.setCellValue(header[k]);
            cell.setCellStyle(styles.get("header"));
            rownum++;
            row = sheet.createRow(rownum);
        }
        rownum = beginROW++;
        cellnum++;
        row = sheet.getRow(rownum);
        for (int k = 0; k < header.length; k++) {
            Cell cell = row.createCell(cellnum);
            cell.setCellValue(header[k]);
            cell.setCellStyle(styles.get("cell_normal_centered"));
            sheet.autoSizeColumn(k);
            rownum++;
            row = sheet.getRow(rownum);
        }
        return sheet;
    }

    public static boolean save(Workbook workb, String filepath) throws IOException {
        if (filepath == null) {
            return false;
        }
        if (workb instanceof HSSFWorkbook) {
            if (!filepath.endsWith("xls")) {
                filepath += ".xls";
            }
        } else if (workb instanceof XSSFWorkbook) {
            if (!filepath.endsWith("xlsx")) {
                filepath += ".xlsx";
            }
        }
        try (FileOutputStream out = new FileOutputStream(new File(filepath))) {
            workb.write(out);
        }
        return true;
    }

    public Map<String, CellStyle> getStyles() {
        return styles;
    }

    public void setStyles(Map<String, CellStyle> styles) {
        this.styles = styles;
    }

    public Workbook getWorkbook() {
        return workbook;
    }

    public void setWorkbook(Workbook workbook) {
        this.workbook = workbook;
    }

    static {
        createStyles(new HSSFWorkbook());
    }
}

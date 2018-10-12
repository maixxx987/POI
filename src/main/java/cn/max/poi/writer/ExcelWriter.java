package cn.max.poi.writer;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * @author MaxStar
 * @date 2018/9/6
 */
public abstract class ExcelWriter {

    /**
     * 输出工作簿
     * 先判断是不是使用SXSSF模式，然后判断是否Excel2003
     * 经测试，自动对齐会大幅影响速度，所以暂时注释掉
     *
     * @param sheetTitle         单元格名称
     * @param headers            表头
     * @param rowList            数据
     * @param isSXSSF            是否使用SXSSF模式(大数据量的时候使用,只支持XLSX)
     * @param rowAcessWindowSize 内存中存储多少列（默认1000列）
     * @param isExcel2003        是否excel2003
     *                           true:HSSF
     *                           false:XSSF
     */
    public static void exportWorkBook(String sheetTitle,
                                      String[] headers,
                                      List<List<String>> rowList,
                                      boolean isSXSSF,
                                      Integer rowAcessWindowSize,
                                      boolean isExcel2003,
                                      String path) throws IOException {
        OutputStream outputStream = new FileOutputStream(path);
        if (isSXSSF) {
            SXSSFWorkbook sxssfWorkbook;
            if (rowAcessWindowSize == null) {
                sxssfWorkbook = new SXSSFWorkbook(1000);
            } else {
                sxssfWorkbook = new SXSSFWorkbook(rowAcessWindowSize);
            }
            SXSSFSheet sheet = sxssfWorkbook.createSheet(sheetTitle);
//            sheet.trackAllColumnsForAutoSizing();
//            createHeader(headers, sheet, this.setHeaderStyle(sxssfWorkbook));
//            createBody(rowList, sheet, this.setBodyStyle(sxssfWorkbook));
            createHeader(headers, sheet, setHeaderStyle(sxssfWorkbook));
            createBody(rowList, sheet, setBodyStyle(sxssfWorkbook));
            sxssfWorkbook.write(outputStream);
            outputStream.close();
            sxssfWorkbook.dispose();
        } else {
            Workbook workBook;
            if (isExcel2003) {
                workBook = new HSSFWorkbook();
            } else {
                workBook = new XSSFWorkbook();
            }
            Sheet sheet = workBook.createSheet(sheetTitle);
            createHeader(headers, sheet, setHeaderStyle(workBook));
            createBody(rowList, sheet, setBodyStyle(workBook));
            workBook.write(outputStream);
            outputStream.close();
            workBook.close();
        }
    }

    /**
     * 创建表头
     *
     * @param headers 表头
     * @param sheet   表
     */
    private static void createHeader(String[] headers, Sheet sheet, CellStyle headerStyle) {
        Row row = sheet.createRow(0);
        for (int i = 0; i < headers.length; i++) {
            Cell cell = row.createCell(i);
            cell.setCellValue(headers[i]);
            cell.setCellStyle(headerStyle);
        }
    }

    /**
     * 创建正文单元格
     *
     * @param rowList 数据
     * @param sheet   表
     */
    private static void createBody(List<List<String>> rowList, Sheet sheet, CellStyle bodyStyle) {
        for (int i = 0; i < rowList.size(); i++) {
            Row row = sheet.createRow(i + 1);
            List<String> dataList = rowList.get(i);
            for (int j = 0; j < dataList.size(); j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(dataList.get(j));
                cell.setCellStyle(bodyStyle);
            }
//            sheet.autoSizeColumn(i);
        }
    }

    /**
     * 设置表头格式
     *
     * @param workBook 工作簿
     * @return 表头样式
     */
//    abstract CellStyle setHeaderStyle(Workbook workBook);
    private static CellStyle setHeaderStyle(Workbook workBook) {
        CellStyle headerStyle = workBook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        Font font = workBook.createFont();
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        headerStyle.setFont(font);
        return headerStyle;
    }

    /**
     * 设置正文单元格格式
     *
     * @param workBook 工作簿
     * @return 单元格样式
     */
//    abstract CellStyle setBodyStyle(Workbook workBook);
    private static CellStyle setBodyStyle(Workbook workBook) {
        CellStyle bodyStyle = workBook.createCellStyle();
        bodyStyle.setAlignment(HorizontalAlignment.CENTER);
        bodyStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        Font font = workBook.createFont();
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 11);
        bodyStyle.setFont(font);
        return bodyStyle;
    }
}
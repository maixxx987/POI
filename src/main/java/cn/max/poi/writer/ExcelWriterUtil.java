package cn.max.poi.writer;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

/**
 * 写入Excel帮助类
 *
 * @author MaxStar
 * @date 2018/9/6
 */
public class ExcelWriterUtil {

    /**
     * 创建表头
     *
     * @param headers 表头
     * @param sheet   表
     */
    public static void createHeader(String[] headers, Sheet sheet, CellStyle headerStyle) {
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
    public static void createBody(List<List<String>> rowList, Sheet sheet, CellStyle bodyStyle) {
        for (int i = 0; i < rowList.size(); i++) {
            Row row = sheet.createRow(i + 1);
            List<String> dataList = rowList.get(i);
            for (int j = 0; j < dataList.size(); j++) {
                Cell cell = row.createCell(j);
                cell.setCellValue(dataList.get(j));
                cell.setCellStyle(bodyStyle);
            }
            sheet.autoSizeColumn(i);
        }
    }

    /**
     * 输出
     *
     * @param workbook 工作簿
     * @param path     输出路径
     * @throws IOException
     */
    private static void export(Workbook workbook, String path) throws IOException {
        OutputStream outputStream = new FileOutputStream(path);
        workbook.write(outputStream);
        outputStream.close();
    }

    /**
     * 输出
     *
     * @param sxssfWorkbook sxssf工作簿
     * @param path          输出路径
     * @throws IOException
     */
    public static void exportSXSSFWorkbook(SXSSFWorkbook sxssfWorkbook, String path) throws IOException {
        export(sxssfWorkbook, path);
        sxssfWorkbook.dispose();
        sxssfWorkbook.close();
    }

    /**
     * 输出
     *
     * @param workbook 工作簿
     * @param path     输出路径
     * @throws IOException
     */
    public static void exportWorkbook(Workbook workbook, String path) throws IOException {
        export(workbook, path);
        workbook.close();
    }
}
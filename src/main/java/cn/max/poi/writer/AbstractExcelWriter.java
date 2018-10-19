package cn.max.poi.writer;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.util.List;
import java.util.Map;

import static cn.max.poi.writer.ExcelWriterUtil.*;

/**
 * ExcelWriter基类
 *
 * @author MaxStar
 * @date 2018/10/12
 */
public abstract class AbstractExcelWriter {

    private Workbook workbook = null;
    private SXSSFWorkbook sxssfWorkbook = null;


    /**
     * 创建普通工作簿
     *
     * @param isExcel2003 true  -> Excel2003(.xls)
     *                    false -> Excel2007(.xlsx)
     */
    public void createWorkbook(boolean isExcel2003) {
        if (isExcel2003) {
            workbook = new HSSFWorkbook();
        } else {
            workbook = new XSSFWorkbook();
        }
    }

    /**
     * 创建SXSSFWorkbook
     *
     * @param rowAcessWindowSize 内存内的工作簿数据条数
     */
    public void createSXSSFWorkbook(Integer rowAcessWindowSize) {
        if (rowAcessWindowSize == null) {
            sxssfWorkbook = new SXSSFWorkbook(1000);
        } else {
            sxssfWorkbook = new SXSSFWorkbook(rowAcessWindowSize);
        }
    }

    /**
     * 设置表头格式(派生类重写)
     *
     * @param workbook 工作簿
     * @return 表头样式
     */
    abstract CellStyle setHeaderStyle(Workbook workbook);


    /**
     * 设置正文单元格格式(派生类重写)
     *
     * @param workbook 工作簿
     * @return 单元格样式
     */
    abstract CellStyle setBodyStyle(Workbook workbook);


    /**
     * 创建单个工作表的普通工作簿
     *
     * @param sheetTitle 单元表名
     * @param header     表头
     * @param rowList    每一行数据集合
     * @param exportPath 输出目录
     * @throws IOException
     */
    public void writeWorkbook(
            String sheetTitle,
            String[] header,
            List<List<String>> rowList,
            String exportPath) throws IOException {
        Sheet sheet = workbook.createSheet(sheetTitle);
        createHeader(header, sheet, this.setHeaderStyle(workbook));
        createBody(rowList, sheet, this.setBodyStyle(workbook));
        exportWorkbook(workbook, exportPath);
    }

    /**
     * 创建多个工作表的普通工作簿
     *
     * @param sheetTitles 表名集合
     * @param headers     表头集合
     * @param sheetData   每个表的数据集合
     * @param exportPath  输出路径
     * @throws IOException
     */
    public void writeWorkbookMultiSheet(
            String[] sheetTitles,
            String[][] headers,
            Map<String, List<List<String>>> sheetData,
            String exportPath) throws IOException {
        for (String sheetTitle : sheetTitles) {
            Sheet sheet = workbook.createSheet(sheetTitle);
            for (String[] header : headers) {
                createHeader(header, sheet, this.setHeaderStyle(workbook));
                createBody(sheetData.get(sheetTitle), sheet, this.setBodyStyle(workbook));
            }
        }
        exportWorkbook(workbook, exportPath);
    }

    /**
     * 创建单个工作表的SXSSF工作簿
     *
     * @param sheetTitle 单元表名
     * @param header     表头
     * @param rowList    每一行数据集合
     * @param exportPath 输出目录
     * @throws IOException
     */
    public void writeSXSSFWorkbook(
            String sheetTitle,
            String[] header,
            List<List<String>> rowList,
            String exportPath) throws IOException {
        SXSSFSheet sheet = sxssfWorkbook.createSheet(sheetTitle);
        sheet.trackAllColumnsForAutoSizing();
        createHeader(header, sheet, this.setHeaderStyle(sxssfWorkbook));
        createBody(rowList, sheet, this.setBodyStyle(sxssfWorkbook));
        exportSXSSFWorkbook(sxssfWorkbook, exportPath);
    }

    /**
     * 创建多个工作表的SXSSF工作簿
     *
     * @param sheetTitles 表名集合
     * @param headers     表头集合
     * @param sheetData   每个表的数据集合
     * @param exportPath  输出路径
     * @throws IOException
     */
    public void writeSXSSFWorkbookMultiSheet(
            String[] sheetTitles,
            String[][] headers,
            Map<String, List<List<String>>> sheetData,
            String exportPath) throws IOException {
        for (String sheetTitle : sheetTitles) {
            SXSSFSheet sheet = sxssfWorkbook.createSheet(sheetTitle);
            sheet.trackAllColumnsForAutoSizing();
            for (String[] header : headers) {
                createHeader(header, sheet, this.setHeaderStyle(sxssfWorkbook));
                createBody(sheetData.get(sheetTitle), sheet, this.setBodyStyle(sxssfWorkbook));
            }
        }
        exportSXSSFWorkbook(sxssfWorkbook, exportPath);
    }
}
package cn.max.poi.reader;

import cn.max.poi.reader.ExcelReader;
import cn.max.poi.reader.ExcelReader2003;
import cn.max.poi.reader.ExcelReader2007;
import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

/**
 * @author MaxStar
 * @date 2018/8/31
 */
public class ExcelReaderTest {
    private InputStream inputStream;

    @Before
    public void setUp() throws FileNotFoundException {
        inputStream = new FileInputStream(new File(this.getClass().getResource("/data.xls").getFile()));
        inputStream = new FileInputStream(new File(this.getClass().getResource("/data.xlsx").getFile()));
    }

    /**
     * 解析单个sheet
     *
     * @throws Exception
     */
    @Test
    public void testResolve() throws Exception {

        long start = System.currentTimeMillis();

        /*****  ExcelReader  *****/
//        List<List<String>> rowList = ExcelReader.process(inputStream, "Sheet1");

        /*****  ExcelReader2003  *****/
//        ExcelReader2003 reader = new ExcelReader2003();
//        reader.process(inputStream, "Sheet1", true);
//        List<List<String>> rowList = reader.getRowList();

        /*****  ExcelReader2007  *****/
        ExcelReader2007 reader = new ExcelReader2007();
        reader.process(inputStream, "Sheet1");
        List<List<String>> rowList = reader.getRowList();

        long end = System.currentTimeMillis() - start;

        // 遍历单元格内容
        rowList.forEach(rowValueList -> {
                    rowValueList.forEach(rowValue -> System.out.print(rowValue + " "));
                    System.out.println();
                }
        );

        System.out.println("当前sheet行数：" + rowList.size());
        System.out.println("耗时：" + end + "ms");
    }

    /**
     * 解析多个sheet
     *
     * @throws Exception
     */
    @Test
    public void testResolveAll() throws Exception {
        long start = System.currentTimeMillis();

        /*****  ExcelReader  *****/
        Map<String, List<List<String>>> sheetMap = ExcelReader.processAll(inputStream);

        /*****  ExcelReader2003  *****/
//        ExcelReader2003 reader = new ExcelReader2003();
//        reader.process(inputStream, "Sheet1", false);
//        Map<String, List<List<String>>> sheetMap  = reader.getSheetMap();

        /*****  ExcelReader2007  *****/
//        ExcelReader2007 reader = new ExcelReader2007();
//        reader.processAll(inputStream);
//        Map<String, List<List<String>>> sheetMap = reader.getSheetMap();

        long end = System.currentTimeMillis() - start;

        // 遍历单元格内容
        sheetMap.forEach((name, rowList) -> {
            System.out.println("正在解析:" + name);
            rowList.forEach(rowValueList -> {
                rowValueList.forEach(rowValue -> System.out.print(rowValue + " "));
                System.out.println();
            });
            System.out.println(name + "解析结束");
            System.out.println("===========================================");
        });
        System.out.println("总耗时：" + end);
    }
}

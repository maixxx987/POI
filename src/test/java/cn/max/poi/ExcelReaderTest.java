package cn.max.poi;

import cn.max.poi.reader.ExcelReader;
import cn.max.poi.reader.ExcelReader2007;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.InputStream;
import java.util.List;

/**
 * @author MaxStar
 * @date 2018/8/31
 */
public class ExcelReaderTest {
    private InputStream inputStream;

    /**
     * 解析单个sheet
     *
     * @throws Exception
     */
    @Test
    public void testResolve() throws Exception {
//        inputStream = new FileInputStream(new File(this.getClass().getResource("/data.xls").getFile()));
        inputStream = new FileInputStream(new File(this.getClass().getResource("/data.xlsx").getFile()));

        long start = System.currentTimeMillis();

        /*****  ExcelReader  *****/
        List<List<String>> rowList = ExcelReader.process(inputStream, "Sheet1");

        /*****  ExcelReader2003  *****/
//        ExcelReader2003 reader = new ExcelReader2003();
//        reader.process(inputStream, "Sheet1");
//        List<List<String>> rowList = reader.getRowList();

        /*****  ExcelReader2007  *****/
//        ExcelReader2007 reader = new ExcelReader2007();
//        reader.process(inputStream, "Sheet1");
//        List<List<String>> rowList = reader.getRowList();

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
        List<List<List<String>>> sheetList = ExcelReader.processAll(inputStream);


        /*****  ExcelReader2007  *****/
//        ExcelReader2007 reader = new ExcelReader2007();
//        reader.process(inputStream, "Sheet1");
//        List<List<List<String>>> sheetList = reader.getSheetList();

        long end = System.currentTimeMillis() - start;

        // 遍历单元格内容
        for (int i = 0; i < sheetList.size(); i++) {
            List<List<String>> rowList = sheetList.get(i);
            System.out.println("遍历第" + (i + 1) + "个sheet的内容");
            rowList.forEach(rowValueList -> {
                        rowValueList.forEach(rowValue -> System.out.print(rowValue + " "));
                        System.out.println();
                    }
            );
            System.out.println("第" + (i + 1) + "个sheet遍历完毕，行数：" + rowList.size());
        }
        System.out.println("总耗时：" + end);
    }
}

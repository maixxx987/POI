package cn.max.poi;

import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;

/**
 * @author MaxStar
 * @date 2018/8/7
 */
public class ExcelReaderTest {

    private ExcelReader reader;
    private InputStream inputStream;
    private File file;
    private String filePath;

    @Before
    public void setUp() throws FileNotFoundException {
        reader = new ExcelReader();
        // 文件
        file = new File(this.getClass().getResource("/data.xlsx").getFile());
        // 输入流
        //inputStream = new FileInputStream(file);
        // 文件路径
        //filePath = this.getClass().getResource("/data.xlsx").getPath();
    }

    /**
     * 解析单个sheet
     *
     * @throws Exception
     */
    @Test
    public void testResolveOne() throws Exception {
        long star = System.currentTimeMillis();
        reader.processOne(file, 1);
//        reader.processOne(inputStream, 1);
//        reader.processOne(filePath, 1);
        List<List<String>> rowList = reader.getRowList();
        long end = System.currentTimeMillis() - star;

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
        long star = System.currentTimeMillis();
        reader.processAll(file);
//        reader.processAll(inputStream);
//        reader.processAll(filePath);
        List<List<List<String>>> sheetList = reader.getSheetList();
        long end = System.currentTimeMillis() - star;

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

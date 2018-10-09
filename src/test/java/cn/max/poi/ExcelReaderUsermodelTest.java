package cn.max.poi;

import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;

/**
 * @author MaxStar
 * @date 2018/8/31
 */
public class ExcelReaderUsermodelTest {
    private ExcelReaderUsermodel reader;
    private InputStream inputStream;

    @Before
    public void setUp() throws FileNotFoundException {
        reader = new ExcelReaderUsermodel();
//        inputStream = new FileInputStream(new File(this.getClass().getResource("/data.xls").getFile()));
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
        List<List<String>> rowList = reader.process(inputStream, "Sheet1");
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
}

package cn.max.poi;

import org.junit.Before;
import org.junit.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;

/**
 * 事件驱动(SAX)解析 Excel97-2003
 *
 * @author MaxStar
 * @date 2018/8/7
 */
public class ExcelReader2003Test {

    private ExcelReader2003 reader;
    private InputStream inputStream;

    @Before
    public void setUp() throws FileNotFoundException {
        reader = new ExcelReader2003();
        inputStream = new FileInputStream(new File(this.getClass().getResource("/data.xls").getFile()));
    }

    /**
     * 解析单个sheet
     *
     * @throws Exception
     */
    @Test
    public void testResolve() throws Exception {
        long start = System.currentTimeMillis();
        reader.process(inputStream, "Sheet3");
        List<List<String>> rowList = reader.getRowValueList();
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

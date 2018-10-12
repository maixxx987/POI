package cn.max.poi.writer;

import com.google.common.collect.Lists;
import org.apache.poi.ss.usermodel.*;
import org.junit.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.List;

/**
 * @author MaxStar
 * @date 2018/10/11
 */
public class ExcelWriterTest {

    @Test
    public void testGenerateExcel() throws IOException {
        String sheetTitle = "通讯录";
        String[] title = {"姓名", "电话", "生日", "备注"};
        List<String> row1 = Lists.newArrayList("小明", "1234556", "1998-09-07", "BCSASD");
        List<String> row2 = Lists.newArrayList("小黄", "8977902", "1967-12-28", "喜欢车大炮");
        List<String> row3 = Lists.newArrayList("小林", "2344365", "1988-06-24", "有人情味，分得清是非");
        List<List<String>> rowList = new ArrayList<>(100000);
        for (int i = 0; i < 100000; i++) {
            rowList.add(row1);
            rowList.add(row2);
            rowList.add(row3);
        }
        long startTime = System.currentTimeMillis();
        ExcelWriter.exportWorkBook(sheetTitle, title, rowList, true, null, false, "C:\\test\\data.xlsx");
        long endTime = System.currentTimeMillis() - startTime;
        System.out.println("耗时：" + endTime);
    }
}
package cn.max.poi.writer;

import com.google.common.collect.Lists;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Test;

import java.io.File;
import java.io.IOException;
import java.util.List;

/**
 * @author MaxStar
 * @date 2018/10/11
 */
public class ExcelWriterTest {

    @Test
    public void testGenerateExcel() throws IOException {
        String sheetTitle = "通讯录";
        String[] title = {"姓名", "电话", "备注"};
        List<String> row1 = Lists.newArrayList("小明", "1234556");
        List<String> row2 = Lists.newArrayList("小黄", "8977902");
        List<String> row3 = Lists.newArrayList("小林", "23443645");
        List<List<String>> rowList = Lists.newArrayList(row1, row2, row3);
        Workbook xls = ExcelWriter.createWorkBook(sheetTitle, title, rowList, true);
        ExcelWriter.write("C:\\test\\data.xls", xls);

        Workbook xlsx = ExcelWriter.createWorkBook(sheetTitle, title, rowList, false);
        ExcelWriter.write("C:\\test\\data.xlsx", xlsx);
    }
}
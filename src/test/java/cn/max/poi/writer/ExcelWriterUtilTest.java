package cn.max.poi.writer;

import com.google.common.collect.Lists;
import org.junit.Test;

import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * 将数据写入到Excel
 *
 * @author MaxStar
 * @date 2018/10/11
 */
public class ExcelWriterUtilTest {

    /**
     * 写入方法说明：
     * 1.构建数据（单元表名，表头，表格具体数据）
     * 2.从AbstractExcelWriter派生一个类，重写setHeaderStyle和setBodyStyle，用于写入表头格式和数据格式
     * 3.初始化派生类，生成所需要的workbook，然后输出
     *
     * @throws IOException
     */
    @Test
    public void testCreateExcel() throws IOException {
        // sheet name
        String sheetTitle = "通讯录";
        String[] sheetTitles = {"通讯录", "电话录", "备注"};

        // Excel header
        String[] title = {"姓名", "电话", "生日", "备注"};
        String[][] titles = {{"姓名", "电话", "生日", "备注"}, {"姓名", "电话", "生日", "备注"}, {"姓名", "电话", "生日", "备注"}};

        // row data
        List<String> row1 = Lists.newArrayList("小明", "1234556", "1998-09-07", "BCSASD");
        List<String> row2 = Lists.newArrayList("小黄", "8977902", "1967-12-28", "喜欢车大炮");
        List<String> row3 = Lists.newArrayList("小林", "2344365", "1988-06-24", "有人情味，分得清是非");
        String exportPathXls = "C:\\test\\data3.xls";
        String exportPathXlsx = "C:\\test\\data3.xlsx";

        // sheet data
        List<List<String>> rowList = new ArrayList<>(100000);
        for (int i = 0; i < 100; i++) {
            rowList.add(row1);
            rowList.add(row2);
            rowList.add(row3);
        }

        // sheets
        Map<String, List<List<String>>> map = new HashMap<>();
        map.put(sheetTitles[0], rowList);
        map.put(sheetTitles[1], rowList);
        map.put(sheetTitles[2], rowList);


        long startTime = System.currentTimeMillis();

        // 1.派生类初始化
        MyExcelWriter myExcelWriter = new MyExcelWriter();

        // 2.派生类中创建工作簿
        // SXSSF
        myExcelWriter.createSXSSFWorkbook(1000);

        // 3.输出excel
        myExcelWriter.writeSXSSFWorkbook(sheetTitle, title, rowList, exportPathXlsx);
//        myExcelWriter.writeSXSSFWorkbookMultiSheet(sheetTitles, titles, map, exportPathXlsx);

        // 2003
//        myExcelWriter.createWorkbook(true);
//        myExcelWriter.writeWorkbook(sheetTitle, title, rowList, exportPathXls);
//        myExcelWriter.writeWorkbookMultiSheet(sheetTitles, titles, map, exportPathXls);

        // 2007
//        myExcelWriter.createWorkbook(false);
//        myExcelWriter.writeWorkbook(sheetTitle, title, rowList, exportPathXlsx);
//        myExcelWriter.writeWorkbookMultiSheet(sheetTitles, titles, map, exportPathXlsx);

        long endTime = System.currentTimeMillis() - startTime;
        System.out.println("耗时：" + endTime);
    }
}
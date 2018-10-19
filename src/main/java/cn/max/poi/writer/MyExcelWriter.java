package cn.max.poi.writer;

import org.apache.poi.ss.usermodel.*;

/**
 * AbstractExcelWriter派生类，重写setHeaderStyle和setBodyStyle方法，写入所需要的格式
 *
 * @author MaxStar
 * @date 2018/10/15
 */
public class MyExcelWriter extends AbstractExcelWriter {

    /**
     * 表头格式
     *
     * @param workbook 工作簿
     * @return
     */
    @Override
    CellStyle setHeaderStyle(Workbook workbook) {
        CellStyle headerStyle = workbook.createCellStyle();
        headerStyle.setFillForegroundColor(IndexedColors.LIGHT_YELLOW.index);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        Font font = workbook.createFont();
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 12);
        font.setBold(true);
        headerStyle.setFont(font);
        return headerStyle;
    }

    /**
     * 数据格式
     *
     * @param workbook 工作簿
     * @return
     */
    @Override
    CellStyle setBodyStyle(Workbook workbook) {
        CellStyle bodyStyle = workbook.createCellStyle();
        bodyStyle.setAlignment(HorizontalAlignment.CENTER);
        bodyStyle.setVerticalAlignment(VerticalAlignment.CENTER);

        Font font = workbook.createFont();
        font.setFontName("微软雅黑");
        font.setFontHeightInPoints((short) 11);
        bodyStyle.setFont(font);
        return bodyStyle;
    }
}

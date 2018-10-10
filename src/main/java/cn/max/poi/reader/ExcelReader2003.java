package cn.max.poi.reader;

import org.apache.poi.hssf.eventusermodel.EventWorkbookBuilder.SheetRecordCollectingListener;
import org.apache.poi.hssf.eventusermodel.*;
import org.apache.poi.hssf.eventusermodel.dummyrecord.LastCellOfRowDummyRecord;
import org.apache.poi.hssf.eventusermodel.dummyrecord.MissingCellDummyRecord;
import org.apache.poi.hssf.model.HSSFFormulaParser;
import org.apache.poi.hssf.record.*;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

import static org.apache.commons.lang3.StringUtils.isBlank;

/**
 * 使用SAX方法解析Excel（只能解析2003版本，即尾缀为.xls）
 *
 * @author MaxStar
 * @date 2018/8/21
 */
public class ExcelReader2003 implements HSSFListener {

    /**
     * true获取结果值，false获取公式
     */
    private boolean outputFormulaValues = true;

    /**
     * 需解析的表名
     */
    private String sheetName;

    /**
     * 解析公式
     */
    private SheetRecordCollectingListener workbookBuildingListener;
    private HSSFWorkbook stubWorkbook;

    /**
     * 解析单元格内容
     */
    private SSTRecord sstRecord;
    private FormatTrackingHSSFListener formatListener;

    /**
     * 获取sheet信息
     */
    private int sheetIndex = -1;
    private BoundSheetRecord[] orderedBSRs;
    private List<BoundSheetRecord> boundSheetRecords = new ArrayList<>();

    /**
     * 处理公式字符串
     */
    private boolean outputNextStringRecord;

    /**
     * 判断是否找到需要解析的单元格
     */
    private boolean isGetSheetName = false;

    private List<String> cellValueList = new ArrayList<>();
    private List<List<String>> rowList = new ArrayList<>();

    /**
     * 解析初始化
     */
    public void process(InputStream in, String sheetName) throws IOException {
        POIFSFileSystem fs = new POIFSFileSystem(in);
        this.sheetName = isBlank(sheetName) ? "Sheet1" : sheetName;
        MissingRecordAwareHSSFListener listener = new MissingRecordAwareHSSFListener(this);
        formatListener = new FormatTrackingHSSFListener(listener);

        HSSFEventFactory factory = new HSSFEventFactory();
        HSSFRequest request = new HSSFRequest();

        if (outputFormulaValues) {
            request.addListenerForAllRecords(formatListener);
        } else {
            workbookBuildingListener = new SheetRecordCollectingListener(formatListener);
            request.addListenerForAllRecords(workbookBuildingListener);
        }

        factory.processWorkbookEvents(request, fs);
    }

    /**
     * 解析
     */
    @Override
    public void processRecord(Record record) {
        String thisStr = null;
        short sid = record.getSid();

        // 解析sheet
        if (sid == BoundSheetRecord.sid) {
            boundSheetRecords.add((BoundSheetRecord) record);
        } else if (sid == BOFRecord.sid) {
            BOFRecord br = (BOFRecord) record;
            if (br.getType() == BOFRecord.TYPE_WORKSHEET) {
                // Create sub workbook if required
                if (workbookBuildingListener != null && stubWorkbook == null) {
                    stubWorkbook = workbookBuildingListener.getStubHSSFWorkbook();
                }

                sheetIndex++;
                if (orderedBSRs == null) {
                    orderedBSRs = BoundSheetRecord.orderByBofPosition(boundSheetRecords);
                }

                // 判断是否与需要解析的sheet名是否相同
                isGetSheetName = orderedBSRs[sheetIndex].getSheetname().trim().equals(sheetName);
            }
        } else if (sid == SSTRecord.sid) {
            sstRecord = (SSTRecord) record;
        } else {
            if (isGetSheetName) {
                switch (sid) {
                    // 空单元格
                    case BlankRecord.sid:
                        cellValueList.add(null);
                        break;

                    // 公式
                    case FormulaRecord.sid:
                        FormulaRecord frec = (FormulaRecord) record;

                        // true解析公式，false获取值
                        if (outputFormulaValues) {
                            if (Double.isNaN(frec.getValue())) {
                                outputNextStringRecord = true;
                            } else {
                                thisStr = formatListener.formatNumberDateCell(frec);
                            }
                        } else {
                            thisStr = HSSFFormulaParser.toFormulaString(stubWorkbook, frec.getParsedExpression());
                        }
                        break;

                    // 公式字符串
                    case StringRecord.sid:
                        if (outputNextStringRecord) {
                            // String for formula
                            StringRecord srec = (StringRecord) record;
                            thisStr = srec.getString().trim();
                            outputNextStringRecord = false;
                        }
                        break;
                    case LabelRecord.sid:
                        LabelRecord lrec = (LabelRecord) record;
                        thisStr = lrec.getValue().trim();
                        break;

                    // 字符串单元格
                    case LabelSSTRecord.sid:
                        LabelSSTRecord lsrec = (LabelSSTRecord) record;
                        if (sstRecord == null) {
                            cellValueList.add(null);
                        } else {
                            thisStr = sstRecord.getString(lsrec.getSSTIndex()).toString().trim();
                        }
                        break;

                    // 数字/日期/货币/科学计数 单元格
                    case NumberRecord.sid:
                        NumberRecord numrec = (NumberRecord) record;
                        thisStr = formatListener.formatNumberDateCell(numrec);
                        break;
                    default:
                        break;
                }

                if (record instanceof MissingCellDummyRecord) {
                    cellValueList.add(null);
                }

                // 赋值
                if (thisStr != null) {
                    if (thisStr.equals("")) {
                        thisStr = null;
                    }
                    cellValueList.add(thisStr);
                }

                // 行尾
                if (record instanceof LastCellOfRowDummyRecord) {
                    rowList.add(cellValueList);
                    cellValueList = new ArrayList<>();
                }
            }
        }
    }

    public List<List<String>> getRowList() {
        return rowList;
    }
}
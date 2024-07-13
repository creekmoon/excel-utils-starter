package cn.creekmoon.excel.core.R.reader.cell;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.StringConverter;
import cn.creekmoon.excel.core.R.reader.ReaderContext;
import cn.creekmoon.excel.core.R.readerResult.ReaderResult;
import cn.creekmoon.excel.core.R.readerResult.cell.CellReaderResult;
import cn.creekmoon.excel.util.ExcelCellUtils;
import cn.creekmoon.excel.util.exception.CheckedExcelException;
import cn.creekmoon.excel.util.exception.ExConsumer;
import cn.creekmoon.excel.util.exception.ExFunction;
import cn.creekmoon.excel.util.exception.GlobalExceptionMsgManager;
import cn.hutool.core.text.StrFormatter;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.sax.Excel07SaxReader;
import cn.hutool.poi.excel.sax.handler.RowHandler;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;

import java.io.IOException;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.function.BiConsumer;

import static cn.creekmoon.excel.util.ExcelConstants.CONVERT_FAIL_MSG;
import static cn.creekmoon.excel.util.ExcelConstants.FIELD_LACK_MSG;

@Slf4j
public class HutoolCellReader<R> implements CellReader<R> {


    protected ExcelImport parent;


    public HutoolCellReader(ExcelImport parent) {
        this.parent = parent;
    }


    /**
     * 获取SHEET页的总行数
     *
     * @return
     */
    @SneakyThrows
    public Long getSheetRowCount() {
        return getExcelImport().getSheetRowCount(getReaderContext().sheetIndex);
    }

    @Override
    public <T> HutoolCellReader<R> addConvert(String cellReference, ExFunction<String, T> convert, BiConsumer<R, T> setter) {

        return addConvert(ExcelCellUtils.excelCellToRowIndex(cellReference), ExcelCellUtils.excelCellToColumnIndex(cellReference), convert, setter);
    }

    @Override
    public HutoolCellReader<R> addConvert(String cellReference, BiConsumer<R, String> reader) {
        return addConvert(ExcelCellUtils.excelCellToRowIndex(cellReference), ExcelCellUtils.excelCellToColumnIndex(cellReference), x -> x, reader);
    }

    @Override
    public <T> HutoolCellReader<R> addConvert(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        getReaderContext().cell2converts.computeIfAbsent(rowIndex, HashMap::new);
        getReaderContext().cell2converts.get(rowIndex).put(colIndex, convert);
        getReaderContext().cell2consumers.computeIfAbsent(rowIndex, HashMap::new);
        getReaderContext().cell2consumers.get(rowIndex).put(colIndex, setter);
        return this;
    }

    @Override
    public HutoolCellReader<R> addConvert(int rowIndex, int colIndex, BiConsumer<R, String> setter) {
        return addConvert(rowIndex, colIndex, x -> x, setter);
    }


    @Override
    public <T> HutoolCellReader<R> addConvertAndSkipEmpty(int rowIndex, int colIndex, BiConsumer<R, String> setter) {
        return addConvertAndSkipEmpty(rowIndex, colIndex, x -> x, setter);
    }

    @Override
    public <T> HutoolCellReader<R> addConvertAndSkipEmpty(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        getReaderContext().skipEmptyCells.computeIfAbsent(rowIndex, HashSet::new);
        getReaderContext().skipEmptyCells.get(rowIndex).add(colIndex);
        return addConvert(rowIndex, colIndex, convert, setter);
    }

    @Override
    public <T> HutoolCellReader<R> addConvertAndSkipEmpty(String cellReference, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        return addConvertAndSkipEmpty(ExcelCellUtils.excelCellToRowIndex(cellReference), ExcelCellUtils.excelCellToColumnIndex(cellReference), convert, setter);
    }

    @Override
    public HutoolCellReader<R> addConvertAndSkipEmpty(String cellReference, BiConsumer<R, String> setter) {
        return addConvertAndSkipEmpty(cellReference, x -> x, setter);
    }

    @Override
    public <T> HutoolCellReader<R> addConvertAndMustExist(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        getReaderContext().mustExistCells.computeIfAbsent(rowIndex, HashSet::new);
        getReaderContext().mustExistCells.get(rowIndex).add(colIndex);
        return addConvert(rowIndex, colIndex, convert, setter);
    }


    @Override
    public HutoolCellReader<R> addConvertAndMustExist(int rowIndex, int colIndex, BiConsumer<R, String> setter) {
        return addConvertAndMustExist(rowIndex, colIndex, x -> x, setter);
    }

    @Override
    public HutoolCellReader<R> addConvertAndMustExist(String cellReference, BiConsumer<R, String> setter) {
        return addConvertAndMustExist(ExcelCellUtils.excelCellToRowIndex(cellReference), ExcelCellUtils.excelCellToColumnIndex(cellReference), setter);
    }


    /**
     * 添加校验阶段后置处理器 当所有的convert执行完成后会执行这个操作做最后的校验处理
     *
     * @param postProcessor 后置处理器
     * @param <T>
     * @return
     */
    @Override
    public <T> HutoolCellReader<R> addConvertPostProcessor(ExConsumer<R> postProcessor) {
        if (postProcessor != null) {
            this.getReaderContext().cellConvertPostProcessors.add(postProcessor);
        }
        return this;
    }

    @Override
    public ReaderResult read(ExConsumer<R> consumer) throws IOException {
        //新版读取 使用SAX读取模式
        Excel07SaxReader excel07SaxReader = initSaxReader(consumer);
        /*第一个参数 文件流  第二个参数 -1就是读取所有的sheet页*/
        excel07SaxReader.read(this.getExcelImport().sourceFile.getInputStream(), -1);
        return getReadResult();
    }


    @Override
    public ExcelImport getExcelImport() {
        return parent;
    }

    @Override
    public CellReaderResult<R> getReadResult() {
        return (CellReaderResult) parent.sheetIndex2ReadResult.get(getSheetIndex());
    }


    /**
     * 初始化SAX读取器
     *
     * 
     * @param consumer
     * @retur
     */
    Excel07SaxReader initSaxReader(ExConsumer<R> consumer) {
        Integer targetSheetIndex = getReaderContext().sheetIndex;
        getReaderContext().currentNewObject = getReaderContext().newObjectSupplier.get();

        /*返回一个Sax读取器*/
        return new Excel07SaxReader(new RowHandler() {
            int currentSheetIndex = 0;

            @Override
            public void doAfterAllAnalysed() {

                if (targetSheetIndex != currentSheetIndex) {
                    return;
                }

                /*sheet读取结束时*/
                for (ExConsumer convertPostProcessor : getReaderContext().cellConvertPostProcessors) {
                    if (getReaderContext().errorReport.length() > 0) {
                        throw new RuntimeException("导入失败!");
                    }
                    try {
                        convertPostProcessor.accept(getReaderContext().currentNewObject);
                    } catch (Exception e) {
                        getReadResult().errorCount.incrementAndGet();
                        getReaderContext().errorReport.append(GlobalExceptionMsgManager.getExceptionMsg(e));
                    }
                }

                try {
                    consumer.accept((R) getReaderContext().currentNewObject);
                } catch (Exception e) {
                    getReadResult().errorCount.incrementAndGet();
                    getReaderContext().errorReport.append(GlobalExceptionMsgManager.getExceptionMsg(e));
                }
            }


            @Override
            public void handle(int sheetIndex, long rowIndex, List<Object> rowList) {

            }

            @Override
            public void handleCell(int sheetIndex, long rowIndex, int cellIndex, Object value, CellStyle xssfCellStyle) {
                currentSheetIndex = sheetIndex;

                int colIndex = cellIndex;

                if (targetSheetIndex != currentSheetIndex) {
                    return;
                }

                /*解析单个单元格*/
                if (getReaderContext().cell2consumers.size() <= 0
                        || !getReaderContext().cell2consumers.containsKey((int) rowIndex)
                        || !getReaderContext().cell2consumers.get((int) rowIndex).containsKey(colIndex)
                ) {
                    return;
                }

                try {
                    ExFunction cellConverter = getReaderContext().cell2converts.get((int) rowIndex).get(colIndex);
                    BiConsumer cellConsumer = getReaderContext().cell2consumers.get((int) rowIndex).get(colIndex);
                    String cellValue = StringConverter.parse(value);
                    /*检查必填项/检查可填项*/
                    if (StrUtil.isBlank(cellValue)) {
                        if (getReaderContext().mustExistCells.containsKey((int) rowIndex)
                                && getReaderContext().mustExistCells.get((int) rowIndex).contains(colIndex)) {
                            throw new CheckedExcelException(StrFormatter.format(FIELD_LACK_MSG, ExcelCellUtils.excelIndexToCell((int) rowIndex, colIndex)));
                        }
                        if (getReaderContext().skipEmptyCells.containsKey((int) rowIndex)
                                && getReaderContext().skipEmptyCells.get((int) rowIndex).contains(colIndex)) {
                            return;
                        }
                    }
                    Object apply = cellConverter.apply(cellValue);
                    cellConsumer.accept(getReaderContext().currentNewObject, apply);
                } catch (Exception e) {
                    getReadResult().errorCount.incrementAndGet();
                    getReaderContext().errorReport.append(StrFormatter.format(CONVERT_FAIL_MSG, ExcelCellUtils.excelIndexToCell((int) rowIndex, colIndex)))
                            .append(GlobalExceptionMsgManager.getExceptionMsg(e))
                            .append(";");
                }
            }
        });
    }


    @Override
    public ReaderContext getReaderContext() {
        return getExcelImport().sheetIndex2ReaderContext.get(getSheetIndex());
    }

    @Override
    public Integer getSheetIndex() {
        return getExcelImport().sheetIndex2ReaderBiMap.getKey(this);
    }
}

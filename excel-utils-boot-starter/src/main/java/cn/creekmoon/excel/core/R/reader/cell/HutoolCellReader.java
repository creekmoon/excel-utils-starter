package cn.creekmoon.excel.core.R.reader.cell;

import cn.creekmoon.excel.core.ExcelUtilsConfig;
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
import java.time.LocalDateTime;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Set;
import java.util.function.BiConsumer;

import static cn.creekmoon.excel.util.ExcelConstants.*;

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
        getReaderContext().cell2setter.computeIfAbsent(rowIndex, HashMap::new);
        getReaderContext().cell2setter.get(rowIndex).put(colIndex, setter);
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

    @Override
    public ReaderResult<R> read(ExConsumer<R> consumer) throws Exception {
        return read().consume(consumer);
    }


    @Override
    public ReaderResult<R> read() throws InterruptedException, IOException {

        //新版读取 使用SAX读取模式
        ExcelUtilsConfig.importParallelSemaphore.acquire();
        ((CellReaderResult) getReadResult()).readStartTime = LocalDateTime.now();
        try {
            /*模版一致性检查:  获取声明的所有CELL, 接下来如果读取到cell就会移除, 当所有cell命中时说明单元格是一致的.*/
            Set<String> templateConsistencyCheckCells = new HashSet<>();
            if (getReaderContext().ENABLE_TEMPLATE_CONSISTENCY_CHECK) {
                getReaderContext().cell2setter.forEach((rowIndex, colIndexMap) -> {
                    colIndexMap.forEach((colIndex, var) -> {
                        templateConsistencyCheckCells.add(ExcelCellUtils.excelIndexToCell(rowIndex, colIndex));
                    });
                });
            }

            Excel07SaxReader excel07SaxReader = initSaxReader(templateConsistencyCheckCells);
            /*第一个参数 文件流  第二个参数 -1就是读取所有的sheet页*/
            excel07SaxReader.read(this.getExcelImport().sourceFile.getInputStream(), -1);

            /*模版一致性检查失败*/
            if (getReaderContext().ENABLE_TEMPLATE_CONSISTENCY_CHECK && !templateConsistencyCheckCells.isEmpty()) {
                getReadResult().EXISTS_READ_FAIL.set(true);
                getReadResult().errorCount.incrementAndGet();
                getReadResult().getErrorReport().append(StrFormatter.format(TITLE_CHECK_ERROR));
            }

        } catch (Exception e) {
            log.error("SaxReader读取Excel文件异常", e);
        } finally {
            getReadResult().readSuccessTime = LocalDateTime.now();
            /*释放信号量*/
            ExcelUtilsConfig.importParallelSemaphore.release();
        }
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
     * @param
     * @param templateConsistencyCheckCells
     * @retur
     */
    Excel07SaxReader initSaxReader(Set<String> templateConsistencyCheckCells) {
        Integer targetSheetIndex = getReaderContext().sheetIndex;
        getReaderContext().currentNewObject = getReaderContext().newObjectSupplier.get();


        /*返回一个Sax读取器*/
        return new Excel07SaxReader(new RowHandler() {
            int currentSheetIndex = 0;


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
                if (getReaderContext().cell2setter.size() <= 0
                        || !getReaderContext().cell2setter.containsKey((int) rowIndex)
                        || !getReaderContext().cell2setter.get((int) rowIndex).containsKey(colIndex)
                ) {
                    return;
                }
                /*标题一致性检查*/
                templateConsistencyCheckCells.remove(ExcelCellUtils.excelIndexToCell((int) rowIndex, colIndex));

                try {
                    ExFunction cellConverter = getReaderContext().cell2converts.get((int) rowIndex).get(colIndex);
                    BiConsumer cellConsumer = getReaderContext().cell2setter.get((int) rowIndex).get(colIndex);
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
                    getReadResult().setData((R) getReaderContext().currentNewObject);
                } catch (Exception e) {
                    getReadResult().EXISTS_READ_FAIL.set(true);
                    getReadResult().errorCount.incrementAndGet();
                    getReadResult().getErrorReport().append(StrFormatter.format(CONVERT_FAIL_MSG, ExcelCellUtils.excelIndexToCell((int) rowIndex, colIndex)))
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

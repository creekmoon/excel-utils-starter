package cn.creekmoon.excelUtils.core;

import cn.creekmoon.excelUtils.converter.StringConverter;
import cn.creekmoon.excelUtils.core.reader.ICellReader;
import cn.creekmoon.excelUtils.exception.CheckedExcelException;
import cn.creekmoon.excelUtils.exception.GlobalExceptionManager;
import cn.creekmoon.excelUtils.util.ExcelCellUtils;
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

import static cn.creekmoon.excelUtils.core.ExcelConstants.CONVERT_FAIL_MSG;
import static cn.creekmoon.excelUtils.core.ExcelConstants.FIELD_LACK_MSG;

@Slf4j
public class HutoolCellReader<R> implements ICellReader<R> {

    protected ReaderContext sheetReaderContext;

    protected ExcelImport parent;


    public static <T> HutoolCellReader<T> of(ReaderContext sheetReaderContext, ExcelImport parent)
    {
        HutoolCellReader<T> newInstant = new HutoolCellReader<>();
        newInstant.sheetReaderContext = sheetReaderContext;
        newInstant.parent = parent;
        return newInstant;
    }


    /**
     * 获取SHEET页的总行数
     *
     * @return
     */
    @SneakyThrows
    public Long getSheetRowCount() {
        return getExcelImport().getSheetRowCount(getSheetReaderContext().sheetIndex);
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
        getSheetReaderContext().cell2converts.computeIfAbsent(rowIndex, HashMap::new);
        getSheetReaderContext().cell2converts.get(rowIndex).put(colIndex, convert);
        getSheetReaderContext().cell2consumers.computeIfAbsent(rowIndex, HashMap::new);
        getSheetReaderContext().cell2consumers.get(rowIndex).put(colIndex, setter);
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
        getSheetReaderContext().skipEmptyCells.computeIfAbsent(rowIndex, HashSet::new);
        getSheetReaderContext().skipEmptyCells.get(rowIndex).add(colIndex);
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
        getSheetReaderContext().mustExistCells.computeIfAbsent(rowIndex, HashSet::new);
        getSheetReaderContext().mustExistCells.get(rowIndex).add(colIndex);
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
            this.getSheetReaderContext().cellConvertPostProcessors.add(postProcessor);
        }
        return this;
    }

    @Override
    public void read(ExConsumer<R> consumer) throws CheckedExcelException, IOException {

        //新版读取 使用SAX读取模式
        Excel07SaxReader excel07SaxReader = initSaxReader(getSheetReaderContext().sheetIndex, consumer);
        /*第一个参数 文件流  第二个参数 -1就是读取所有的sheet页*/
        excel07SaxReader.read(this.getExcelImport().file.getInputStream(), -1);
        if (getSheetReaderContext().errorReport.length() > 0) {
            throw new CheckedExcelException(getSheetReaderContext().errorReport.toString());
        }


    }

    @Override
    public ReaderContext getSheetReaderContext() {
        return sheetReaderContext;
    }

    @Override
    public ExcelImport getExcelImport() {
        return parent;
    }


    /**
     * 初始化SAX读取器
     *
     * @param targetSheetIndex 读取的sheetIndex下标
     * @param consumer
     * @return
     */
    Excel07SaxReader initSaxReader(int targetSheetIndex, ExConsumer<R> consumer) {

        getSheetReaderContext().currentNewObject = getSheetReaderContext().newObjectSupplier.get();

        /*返回一个Sax读取器*/
        return new Excel07SaxReader(new RowHandler() {
            int currentSheetIndex = 0;

            @Override
            public void doAfterAllAnalysed() {

                if (targetSheetIndex != currentSheetIndex) {
                    return;
                }

                /*sheet读取结束时*/
                for (ExConsumer convertPostProcessor : getSheetReaderContext().cellConvertPostProcessors) {
                    if (getSheetReaderContext().errorReport.length() > 0) {
                        throw new RuntimeException("导入失败!");
                    }
                    try {
                        convertPostProcessor.accept(getSheetReaderContext().currentNewObject);
                    } catch (Exception e) {
                        getExcelImport().getErrorCount().incrementAndGet();
                        getSheetReaderContext().errorReport.append(GlobalExceptionManager.getExceptionMsg(e));
                    }
                }

                try {
                    consumer.accept((R) getSheetReaderContext().currentNewObject);
                } catch (Exception e) {
                    getExcelImport().getErrorCount().incrementAndGet();
                    getSheetReaderContext().errorReport.append(GlobalExceptionManager.getExceptionMsg(e));
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
                if (getSheetReaderContext().cell2consumers.size() <= 0
                        || !getSheetReaderContext().cell2consumers.containsKey((int) rowIndex)
                        || !getSheetReaderContext().cell2consumers.get((int) rowIndex).containsKey(colIndex)
                ) {
                    return;
                }

                try {
                    ExFunction cellConverter = getSheetReaderContext().cell2converts.get((int) rowIndex).get(colIndex);
                    BiConsumer cellConsumer = getSheetReaderContext().cell2consumers.get((int) rowIndex).get(colIndex);
                    String cellValue = StringConverter.parse(value);
                    /*检查必填项/检查可填项*/
                    if (StrUtil.isBlank(cellValue)) {
                        if (getSheetReaderContext().mustExistCells.containsKey((int) rowIndex)
                                && getSheetReaderContext().mustExistCells.get((int) rowIndex).contains(colIndex)) {
                            throw new CheckedExcelException(StrFormatter.format(FIELD_LACK_MSG, ExcelCellUtils.excelIndexToCell((int) rowIndex, colIndex)));
                        }
                        if (getSheetReaderContext().skipEmptyCells.containsKey((int) rowIndex)
                                && getSheetReaderContext().skipEmptyCells.get((int) rowIndex).contains(colIndex)) {
                            return;
                        }
                    }
                    Object apply = cellConverter.apply(cellValue);
                    cellConsumer.accept(getSheetReaderContext().currentNewObject, apply);
                } catch (Exception e) {
                    getExcelImport().getErrorCount().incrementAndGet();
                    getSheetReaderContext().errorReport.append(StrFormatter.format(CONVERT_FAIL_MSG, ExcelCellUtils.excelIndexToCell((int) rowIndex, colIndex)))
                            .append(GlobalExceptionManager.getExceptionMsg(e))
                            .append(";");
                }
            }
        });
    }


    @Override
    public ReaderContext getReaderContext() {
        return sheetReaderContext;
    }
}

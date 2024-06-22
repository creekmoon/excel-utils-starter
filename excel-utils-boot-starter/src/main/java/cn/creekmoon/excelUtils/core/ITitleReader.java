package cn.creekmoon.excelUtils.core;

import lombok.SneakyThrows;

import java.util.List;
import java.util.function.BiConsumer;


/**
 * CellReader
 * 遍历所有行, 并按标题读取
 * 一个sheet页只会有多个结果对象, 每行是一个对象
 *
 * @param <R>
 */
public interface ITitleReader<R> {
    @SneakyThrows
    Long getSheetRowCount();

    <T> ITitleReader<R> addConvert(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    ITitleReader<R> addConvert(String title, BiConsumer<R, String> reader);

    <T> ITitleReader<R> addConvertAndSkipEmpty(String title, BiConsumer<R, String> setter);

    <T> ITitleReader<R> addConvertAndSkipEmpty(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    ITitleReader<R> addConvertAndMustExist(String title, BiConsumer<R, String> setter);

    <T> ITitleReader<R> addConvertAndMustExist(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    <T> ITitleReader<R> addConvertPostProcessor(ExConsumer<R> postProcessor);

    ExcelImport read(ExConsumer<R> dataConsumer);

    TitleReadResult<R> read() throws InterruptedException;

    @SneakyThrows
    List<R> readAll();

    ITitleReader<R> range(int titleRowIndex, int firstDataRowIndex, int lastDataRowIndex);

    ITitleReader<R> range(int startRowIndex, int lastRowIndex);

    ITitleReader<R> range(int startRowIndex);

    ITitleReader<R> disableTitleConsistencyCheck();

    ITitleReader<R> disableBlankRowFilter();

    SheetReaderContext getSheetReaderContext();

    ExcelImport getExcelImport();


}

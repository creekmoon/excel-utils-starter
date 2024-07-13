package cn.creekmoon.excel.core.R.reader.cell;

import cn.creekmoon.excel.core.R.reader.Reader;
import cn.creekmoon.excel.core.R.readerResult.ReaderResult;
import cn.creekmoon.excel.util.exception.CheckedExcelException;
import cn.creekmoon.excel.util.exception.ExConsumer;
import cn.creekmoon.excel.util.exception.ExFunction;

import java.io.IOException;
import java.util.function.BiConsumer;


/**
 * CellReader
 * 遍历所有单元格, 并按单元格读取
 * 一个sheet页只会有一个结果对象
 *
 * @param <R>
 */
public interface CellReader<R> extends Reader<R> {
    <T> CellReader<R> addConvert(String cellReference, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    CellReader<R> addConvert(String cellReference, BiConsumer<R, String> reader);

    <T> CellReader<R> addConvert(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    CellReader<R> addConvert(int rowIndex, int colIndex, BiConsumer<R, String> setter);

    <T> CellReader<R> addConvertAndSkipEmpty(int rowIndex, int colIndex, BiConsumer<R, String> setter);

    <T> CellReader<R> addConvertAndSkipEmpty(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    <T> CellReader<R> addConvertAndSkipEmpty(String cellReference, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    CellReader<R> addConvertAndSkipEmpty(String cellReference, BiConsumer<R, String> setter);

    <T> CellReader<R> addConvertAndMustExist(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    CellReader<R> addConvertAndMustExist(int rowIndex, int colIndex, BiConsumer<R, String> setter);

    CellReader<R> addConvertAndMustExist(String cellReference, BiConsumer<R, String> setter);

    <T> CellReader<R> addConvertPostProcessor(ExConsumer<R> postProcessor);

    ReaderResult read(ExConsumer<R> consumer) throws CheckedExcelException, IOException;
}



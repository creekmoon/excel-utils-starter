package cn.creekmoon.excelUtils.core.reader;

import cn.creekmoon.excelUtils.core.ExConsumer;
import cn.creekmoon.excelUtils.core.ExFunction;
import cn.creekmoon.excelUtils.core.HutoolCellReader;
import cn.creekmoon.excelUtils.exception.CheckedExcelException;

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
    <T> HutoolCellReader<R> addConvert(String cellReference, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    HutoolCellReader<R> addConvert(String cellReference, BiConsumer<R, String> reader);

    <T> HutoolCellReader<R> addConvert(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    HutoolCellReader<R> addConvert(int rowIndex, int colIndex, BiConsumer<R, String> setter);

    <T> HutoolCellReader<R> addConvertAndSkipEmpty(int rowIndex, int colIndex, BiConsumer<R, String> setter);

    <T> HutoolCellReader<R> addConvertAndSkipEmpty(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    <T> HutoolCellReader<R> addConvertAndSkipEmpty(String cellReference, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    HutoolCellReader<R> addConvertAndSkipEmpty(String cellReference, BiConsumer<R, String> setter);

    <T> HutoolCellReader<R> addConvertAndMustExist(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    HutoolCellReader<R> addConvertAndMustExist(int rowIndex, int colIndex, BiConsumer<R, String> setter);

    HutoolCellReader<R> addConvertAndMustExist(String cellReference, BiConsumer<R, String> setter);

    <T> HutoolCellReader<R> addConvertPostProcessor(ExConsumer<R> postProcessor);

    void read(ExConsumer<R> consumer) throws CheckedExcelException, IOException;
}



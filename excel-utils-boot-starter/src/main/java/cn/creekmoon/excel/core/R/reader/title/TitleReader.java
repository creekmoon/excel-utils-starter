package cn.creekmoon.excel.core.R.reader.title;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.reader.Reader;
import cn.creekmoon.excel.core.R.reader.ReaderContext;
import cn.creekmoon.excel.core.R.readerResult.title.TitleReaderResult;
import cn.creekmoon.excel.util.exception.ExConsumer;
import cn.creekmoon.excel.util.exception.ExFunction;
import lombok.SneakyThrows;

import java.util.function.BiConsumer;


/**
 * CellReader
 * 遍历所有行, 并按标题读取
 * 一个sheet页只会有多个结果对象, 每行是一个对象
 *
 * @param <R>
 */
public interface TitleReader<R> extends Reader<R> {
    @SneakyThrows
    Long getSheetRowCount();

    <T> TitleReader<R> addConvert(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    TitleReader<R> addConvert(String title, BiConsumer<R, String> reader);

    <T> TitleReader<R> addConvertAndSkipEmpty(String title, BiConsumer<R, String> setter);

    <T> TitleReader<R> addConvertAndSkipEmpty(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    TitleReader<R> addConvertAndMustExist(String title, BiConsumer<R, String> setter);

    <T> TitleReader<R> addConvertAndMustExist(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    <T> TitleReader<R> addConvertPostProcessor(ExConsumer<R> postProcessor);

    TitleReaderResult<R> read(ExConsumer<R> dataConsumer);

    TitleReaderResult<R> read();

    TitleReader<R> range(int titleRowIndex, int firstDataRowIndex, int lastDataRowIndex);

    TitleReader<R> range(int startRowIndex, int lastRowIndex);

    TitleReader<R> range(int startRowIndex);

    Integer getSheetIndex();

    ReaderContext getReaderContext();

    ExcelImport getExcelImport();


}

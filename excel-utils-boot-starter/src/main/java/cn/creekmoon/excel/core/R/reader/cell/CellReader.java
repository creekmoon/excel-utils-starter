package cn.creekmoon.excel.core.R.reader.cell;

import cn.creekmoon.excel.core.R.reader.Reader;
import cn.creekmoon.excel.util.exception.ExFunction;

import java.util.function.BiConsumer;


/**
 * This interface extends the `Reader` interface and provides methods to add cell converters for reading
 * and extracting data from specific cells in a spreadsheet. It allows converting values from string to
 * specified types and setting the converted values into corresponding fields or properties of an object.
 *
 * @param <R> the type of the object being read and populated with data
 */
public interface CellReader<R> extends Reader<R> {

    /**
     * 添加一个单元格转换器
     *
     * @param cellReference 单元格引用名称,例如 "F2"
     * @param convert       数值类型适配器, 例如 String --> Date
     * @param setter        Setter方法, 例如 setStartDate(Date date)
     * @param <T>
     * @return
     */
    <T> CellReader<R> addConvert(String cellReference, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    /**
     * 添加一个单元格转换器
     *
     * @param cellReference 单元格引用名称,例如 "F2"
     * @param reader Setter方法, 例如 setName(String name)
     * @return
     */
    CellReader<R> addConvert(String cellReference, BiConsumer<R, String> reader);

    /**
     * 添加一个单元格转换器
     *
     * @param rowIndex 行索引
     * @param colIndex 列索引
     * @param convert 数值类型适配器, 例如 String --> Date
     * @param setter Setter方法, 例如 setStartDate(Date date)
     * @param <T>
     * @return
     */
    <T> CellReader<R> addConvert(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    /**
     * 添加一个单元格转换器
     *
     * @param rowIndex 行索引
     * @param colIndex 列索引
     * @param setter Setter方法, 例如 setName(String name)
     * @return
     */
    CellReader<R> addConvert(int rowIndex, int colIndex, BiConsumer<R, String> setter);

    /**
     * 添加一个单元格转换器并跳过空值
     *
     * @param rowIndex 行索引
     * @param colIndex 列索引
     * @param setter Setter方法, 例如 setName(String name)
     * @param <T>
     * @return
     */
    <T> CellReader<R> addConvertAndSkipEmpty(int rowIndex, int colIndex, BiConsumer<R, String> setter);

    /**
     * 添加一个单元格转换器并跳过空值
     *
     * @param rowIndex 行索引
     * @param colIndex 列索引
     * @param convert 数值类型适配器,例如 String --> Date
     * @param setter Setter方法,例如 setStartDate(Date date)
     * @param <T>
     * @return
     */
    <T> CellReader<R> addConvertAndSkipEmpty(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    /**
     * 添加一个单元格转换器并跳过空值
     *
     * @param cellReference 单元格引用名称,例如 "F2"
     * @param convert 数值类型适配器,例如 String --> Date
     * @param setter Setter方法,例如 setStartDate(Date date)
     * @param <T>
     * @return
     */
    <T> CellReader<R> addConvertAndSkipEmpty(String cellReference, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    /**
     * 添加一个单元格转换器并跳过空值
     *
     * @param cellReference 单元格引用名称,例如 "F2"
     * @param setter Setter方法,例如 setName(String name)
     * @return
     */
    CellReader<R> addConvertAndSkipEmpty(String cellReference, BiConsumer<R, String> setter);

    /**
     * 添加一个单元格转换器并要求存在值
     *
     * @param rowIndex 行索引
     * @param colIndex 列索引
     * @param convert 数值类型适配器,例如 String -> Date
     * @param setter Setter方法,例如 setStartDate(Date date)
     * @param <T>
     * @return
     */
    <T> CellReader<R> addConvertAndMustExist(int rowIndex, int colIndex, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    /**
     * 添加一个单元格转换器并要求存在值
     *
     * @param rowIndex 行索引
     * @param colIndex 列索引
     * @param setter Setter方法,例如 setName(String name)
     * @return
     */
    CellReader<R> addConvertAndMustExist(int rowIndex, int colIndex, BiConsumer<R, String> setter);

    /**
     * 添加一个单元格转换器并要求存在值
     *
     * @param cellReference 单元格引用名称,例如 "F2"
     * @param setter Setter方法,例如 setName(String name)
     * @return
     */
    CellReader<R> addConvertAndMustExist(String cellReference, BiConsumer<R, String> setter);

}



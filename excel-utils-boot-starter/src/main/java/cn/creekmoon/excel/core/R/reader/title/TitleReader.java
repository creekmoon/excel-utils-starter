package cn.creekmoon.excel.core.R.reader.title;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.reader.Reader;
import cn.creekmoon.excel.core.R.readerResult.title.TitleReaderResult;
import cn.creekmoon.excel.util.exception.ExConsumer;
import cn.creekmoon.excel.util.exception.ExFunction;

import java.util.*;
import java.util.function.BiConsumer;
import java.util.function.Supplier;


/**
 * CellReader
 * 遍历所有行, 并按标题读取
 * 一个sheet页只会有多个结果对象, 每行是一个对象
 *
 * @param <R>
 */
public abstract class TitleReader<R> implements Reader<R> {

    /**
     * 标题行号 这里是0,意味着第一行是标题
     */
    public int titleRowIndex = 0;

    /**
     * 首行数据行号
     */
    public int firstRowIndex = titleRowIndex + 1;

    /**
     * 末行数据行号
     */
    public int latestRowIndex = Integer.MAX_VALUE;

    public int sheetIndex;
    public Supplier newObjectSupplier;
    public Object currentNewObject;


    public HashMap<Integer, String> colIndex2Title = new HashMap<>();

    /* 必填项过滤  key=rowIndex  value=<colIndex> */
    public LinkedHashMap<Integer, Set<Integer>> mustExistCells = new LinkedHashMap<>(32);
    /* 选填项过滤  key=rowIndex  value=<colIndex> */
    public LinkedHashMap<Integer, Set<Integer>> skipEmptyCells = new LinkedHashMap<>(32);

    /* key=rowIndex  value=<colIndex,Consumer> 单元格转换器*/
    public LinkedHashMap<Integer, HashMap<Integer, ExFunction>> cell2converts = new LinkedHashMap(32);

    /* key=rowIndex  value=<colIndex,Consumer> 单元格消费者(通常是setter方法)*/
    public LinkedHashMap<Integer, HashMap<Integer, BiConsumer>> cell2setter = new LinkedHashMap(32);


    /* key=title  value=执行器 */
    public LinkedHashMap<String, ExFunction> title2converts = new LinkedHashMap(32);
    /* key=title value=消费者(通常是setter方法)*/
    public LinkedHashMap<String, BiConsumer> title2consumers = new LinkedHashMap(32);
    public List<ExConsumer> convertPostProcessors = new ArrayList<>();
    /* key=title */
    public Set<String> mustExistTitles = new HashSet<>(32);
    public Set<String> skipEmptyTitles = new HashSet<>(32);

    /*启用空白行过滤*/
    public boolean ENABLE_BLANK_ROW_FILTER = true;
    /*启用模板一致性检查 为了防止模板导入错误*/
    public boolean ENABLE_TEMPLATE_CONSISTENCY_CHECK = true;
    /*标志位, 模板一致性检查已经失败 */
    public boolean TEMPLATE_CONSISTENCY_CHECK_FAILED = false;

    abstract public Long getSheetRowCount();

    abstract public <T> TitleReader<R> addConvert(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    abstract public TitleReader<R> addConvert(String title, BiConsumer<R, String> reader);

    abstract public <T> TitleReader<R> addConvertAndSkipEmpty(String title, BiConsumer<R, String> setter);

    abstract public <T> TitleReader<R> addConvertAndSkipEmpty(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    abstract public TitleReader<R> addConvertAndMustExist(String title, BiConsumer<R, String> setter);

    abstract public <T> TitleReader<R> addConvertAndMustExist(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter);

    abstract public <T> TitleReader<R> addConvertPostProcessor(ExConsumer<R> postProcessor);

    abstract public TitleReaderResult<R> read(ExConsumer<R> dataConsumer);

    abstract public TitleReaderResult<R> read();

    abstract public TitleReader<R> range(int titleRowIndex, int firstDataRowIndex, int lastDataRowIndex);

    abstract public TitleReader<R> range(int startRowIndex, int lastRowIndex);

    abstract public TitleReader<R> range(int startRowIndex);

    abstract public Integer getSheetIndex();

    abstract public ExcelImport getExcelImport();


}

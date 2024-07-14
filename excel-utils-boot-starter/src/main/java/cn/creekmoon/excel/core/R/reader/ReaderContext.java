package cn.creekmoon.excel.core.R.reader;

import cn.creekmoon.excel.util.exception.ExConsumer;
import cn.creekmoon.excel.util.exception.ExFunction;

import java.util.*;
import java.util.function.BiConsumer;
import java.util.function.Supplier;

public class ReaderContext {
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

    public ReaderContext(int sheetIndex, Supplier newObjectSupplier) {
        this.sheetIndex = sheetIndex;
        this.newObjectSupplier = newObjectSupplier;
    }


    /**
     * 获取最后一个标题列索引
     * @return
     */
    public Integer getLastTitleColumnIndex() {
        return this.colIndex2Title.keySet().stream().max(Integer::compareTo).get();
    }
}

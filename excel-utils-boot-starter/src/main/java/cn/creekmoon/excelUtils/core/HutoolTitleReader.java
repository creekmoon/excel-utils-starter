package cn.creekmoon.excelUtils.core;

import cn.creekmoon.excelUtils.converter.StringConverter;
import cn.creekmoon.excelUtils.core.reader.TitleReader;
import cn.creekmoon.excelUtils.exception.CheckedExcelException;
import cn.creekmoon.excelUtils.exception.GlobalExceptionManager;
import cn.hutool.core.text.StrFormatter;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.sax.Excel07SaxReader;
import cn.hutool.poi.excel.sax.handler.RowHandler;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;

import java.util.*;
import java.util.function.BiConsumer;

import static cn.creekmoon.excelUtils.core.ExcelConstants.*;

@Slf4j
public class HutoolTitleReader<R> implements TitleReader<R> {

    protected ReaderContext readerContext;

    protected TitleReaderResult<R> titleReaderResult = new TitleReaderResult<>();

    protected ExcelImport parent;


    /**
     * 获取SHEET页的总行数
     *
     * @return
     */
    @SneakyThrows
    @Override
    public Long getSheetRowCount() {
        return getExcelImport().getSheetRowCount(getReaderContext().sheetIndex);
    }

    @Override
    public <T> HutoolTitleReader<R> addConvert(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        getReaderContext().title2converts.put(title, convert);
        getReaderContext().title2consumers.put(title, setter);
        return this;
    }


    @Override
    public HutoolTitleReader<R> addConvert(String title, BiConsumer<R, String> reader) {
        addConvert(title, x -> x, reader);
        return this;
    }


    @Override
    public <T> HutoolTitleReader<R> addConvertAndSkipEmpty(String title, BiConsumer<R, String> setter) {
        getReaderContext().skipEmptyTitles.add(title);
        addConvert(title, x -> x, setter);
        return this;
    }

    @Override
    public <T> HutoolTitleReader<R> addConvertAndSkipEmpty(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        getReaderContext().skipEmptyTitles.add(title);
        addConvert(title, convert, setter);
        return this;
    }


    @Override
    public HutoolTitleReader<R> addConvertAndMustExist(String title, BiConsumer<R, String> setter) {
        getReaderContext().mustExistTitles.add(title);
        addConvert(title, x -> x, setter);
        return this;
    }

    @Override
    public <T> HutoolTitleReader<R> addConvertAndMustExist(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        getReaderContext().mustExistTitles.add(title);
        addConvert(title, convert, setter);
        return this;
    }


    /**
     * 添加校验阶段后置处理器 当所有的convert执行完成后会执行这个操作做最后的校验处理
     *
     * @param postProcessor 后置处理器
     * @param <T>
     * @return
     */
    @Override
    public <T> HutoolTitleReader<R> addConvertPostProcessor(ExConsumer<R> postProcessor) {
        if (postProcessor != null) {
            this.getReaderContext().convertPostProcessors.add(postProcessor);
        }
        return this;
    }


    @Override
    public TitleReaderResult read(ExConsumer<R> dataConsumer) {
        return read().foreachAndConsume(dataConsumer);

    }

    @SneakyThrows
    @Override
    public TitleReaderResult<R> read()  {
        TitleReaderResult<R> readResult = new TitleReaderResult<>();
        parent.sheetIndex2ReadResult.put(getReaderContext().sheetIndex, readResult);

        /*尝试拿锁*/
        ExcelImport.importSemaphore.acquire();
        try {
            //新版读取 使用SAX读取模式
            Excel07SaxReader excel07SaxReader = initSaxReader(getReaderContext().sheetIndex);
            /*第一个参数 文件流  第二个参数 -1就是读取所有的sheet页*/
            excel07SaxReader.read(this.getExcelImport().sourceFile.getInputStream(), -1);
        } catch (Exception e) {
            log.error("SaxReader读取Excel文件异常", e);
        } finally {
            /*释放信号量*/
            ExcelImport.importSemaphore.release();
        }
        return readResult;
    }


    /**
     * 增加读取范围限制
     *
     * @param titleRowIndex    标题所在的行数(下标按照从0开始, 如果是第一行则填0)
     * @param lastDataRowIndex 最后一条数据所在的行数(下标按照从0开始, 如果是第一行则填0)
     * @return
     */
    @Override
    public HutoolTitleReader<R> range(int titleRowIndex, int firstDataRowIndex, int lastDataRowIndex) {
        this.getReaderContext().titleRowIndex = titleRowIndex;
        this.getReaderContext().firstRowIndex = firstDataRowIndex;
        this.getReaderContext().latestRowIndex = lastDataRowIndex;
        return this;
    }

    /**
     * 增加读取范围限制
     *
     * @param startRowIndex 标题所在的行数(下标按照从0开始, 如果是第一行则填0)
     * @param lastRowIndex  最后一条数据所在的行数(下标按照从0开始, 如果是第一行则填0)
     * @return
     */
    @Override
    public HutoolTitleReader<R> range(int startRowIndex, int lastRowIndex) {
        return range(startRowIndex, startRowIndex + 1, lastRowIndex);
    }

    /**
     * 增加读取范围限制
     *
     * @param startRowIndex 起始行下标(从0开始)
     * @return
     */
    @Override
    public HutoolTitleReader<R> range(int startRowIndex) {
        return range(startRowIndex, startRowIndex + 1, Integer.MAX_VALUE);
    }

    /**
     * 行转换
     *
     * @param row 实际上是Map<String, String>对象
     * @throws Exception
     */
    private R rowConvert(Map<String, Object> row) throws Exception {
        /*进行模板一致性检查*/
        if (getReaderContext().ENABLE_TITLE_CHECK) {
            if (getReaderContext().TITLE_CHECK_FAIL_FLAG || !titleConsistencyCheck(getReaderContext().title2converts.keySet(), row.keySet())) {
                getReaderContext().TITLE_CHECK_FAIL_FLAG = true;
                throw new CheckedExcelException(TITLE_CHECK_ERROR);
            }
        }
        getReaderContext().ENABLE_TITLE_CHECK = false;

        /*过滤空白行*/
        if (getReaderContext().ENABLE_BLANK_ROW_FILTER
                && row.values().stream().allMatch(x -> x == null || "".equals(x))
        ) {
            return null;
        }

        /*初始化空对象*/
        R convertObject = (R) this.getReaderContext().newObjectSupplier.get();
        /*最大转换次数*/
        int maxConvertCount = this.getReaderContext().title2consumers.keySet().size();
        /*执行convert*/
        for (Map.Entry<String, Object> entry : row.entrySet()) {
            /*如果包含不支持的标题,  或者已经超过最大次数则不再进行读取*/
            if (!this.getReaderContext().title2consumers.containsKey(entry.getKey()) || maxConvertCount-- <= 0) {
                continue;
            }
            String value = Optional.ofNullable(entry.getValue()).map(x -> (String) x).orElse("");
            /*检查必填项/检查可填项*/
            if (StrUtil.isBlank(value)) {
                if (this.getReaderContext().mustExistTitles.contains(entry.getKey())) {
                    throw new CheckedExcelException(StrFormatter.format(FIELD_LACK_MSG, entry.getKey()));
                }
                if (this.getReaderContext().skipEmptyTitles.contains(entry.getKey())) {
                    continue;
                }
            }
            /*转换数据*/
            try {
                Object convertValue = this.getReaderContext().title2converts.get(entry.getKey()).apply(value);
                this.getReaderContext().title2consumers.get(entry.getKey()).accept(convertObject, convertValue);
            } catch (Exception e) {
                log.warn("EXCEL导入数据转换失败！", e);
                throw new CheckedExcelException(StrFormatter.format(ExcelConstants.CONVERT_FAIL_MSG + GlobalExceptionManager.getExceptionMsg(e), entry.getKey()));
            }
        }
        return convertObject;
    }

    /**
     * 初始化SAX读取器
     *
     * @param targetSheetIndex 读取的sheetIndex下标
     * @return
     */
    Excel07SaxReader initSaxReader(int targetSheetIndex) {


        /*返回一个Sax读取器*/
        return new Excel07SaxReader(new RowHandler() {

            @Override
            public void doAfterAllAnalysed() {
                /*sheet读取结束时*/
            }


            @Override
            public void handle(int sheetIndex, long rowIndex, List<Object> rowList) {
                if (targetSheetIndex != sheetIndex) {
                    return;
                }

                /*读取标题*/
                if (rowIndex == getReaderContext().titleRowIndex) {
                    for (int colIndex = 0; colIndex < rowList.size(); colIndex++) {
                        readerContext.colIndex2Title.put(colIndex, StringConverter.parse(rowList.get(colIndex)));
                    }
                    return;
                }
                /*只读取指定范围的数据 */
                if (rowIndex == (int) getReaderContext().titleRowIndex
                        || rowIndex < getReaderContext().firstRowIndex
                        || rowIndex > getReaderContext().latestRowIndex) {
                    return;
                }
                /*没有添加 convert直接跳过 */
                if (getReaderContext().title2converts.isEmpty()
                        && getReaderContext().title2consumers.isEmpty()
                ) {
                    return;
                }

                /*Excel解析原生的数据*/
                HashMap<String, Object> rowData = new LinkedHashMap<>();
                for (int colIndex = 0; colIndex < rowList.size(); colIndex++) {
                    rowData.put(readerContext.colIndex2Title.get(colIndex), StringConverter.parse(rowList.get(colIndex)));
                }
                titleReaderResult.rowIndex2rawData.put((int) rowIndex, rowData);
                /*转换成业务对象*/
                R currentObject = null;
                try {
                    /*转换*/
                    currentObject = rowConvert(rowData);
                    if (currentObject == null) {
                        return;
                    }
                    /*转换后置处理器*/
                    for (ExConsumer convertPostProcessor : getReaderContext().convertPostProcessors) {
                        convertPostProcessor.accept(currentObject);
                    }
                    titleReaderResult.rowIndex2msg.put((int) rowIndex, CONVERT_SUCCESS_MSG);
                    rowData.put(RESULT_TITLE, CONVERT_SUCCESS_MSG);
                    /*消费*/
                    titleReaderResult.rowIndex2data.put((int) rowIndex, currentObject);
                } catch (Exception e) {
                    getExcelImport().getErrorCount().incrementAndGet();
                    titleReaderResult.errorCount.incrementAndGet();
                    /*写入导出Excel结果*/
                    titleReaderResult.rowIndex2msg.put((int) rowIndex, GlobalExceptionManager.getExceptionMsg(e));
                    rowData.put(RESULT_TITLE, GlobalExceptionManager.getExceptionMsg(e));
                }
                if (currentObject == null && rowData != null) {
                    //假如存在任一数据convert阶段就失败的单, 将打一个标记
                    titleReaderResult.EXISTS_CONVERT_FAIL.set(true);
                }
            }
        });
    }


    /**
     * 标题一致性检查
     *
     * @param targetTitles 我们声明的要拿取的标题
     * @param sourceTitles 传过来的excel文件标题
     * @return
     */
    private Boolean titleConsistencyCheck(Set<String> targetTitles, Set<String> sourceTitles) {
        if (targetTitles.size() > sourceTitles.size()) {
            return false;
        }
        return sourceTitles.containsAll(targetTitles);
    }

    /**
     * 禁用标题一致性检查
     *
     * @return
     */
    public HutoolTitleReader<R> disableTitleConsistencyCheck() {
        this.getReaderContext().ENABLE_TITLE_CHECK = false;
        return this;
    }

    /**
     * 禁用空白行过滤
     *
     * @return
     */
    public HutoolTitleReader<R> disableBlankRowFilter() {
        this.getReaderContext().ENABLE_BLANK_ROW_FILTER = false;
        return this;
    }

    @Override
    public ReaderContext getReaderContext() {
        return readerContext;
    }


    @Override
    public ExcelImport getExcelImport() {
        return getExcelImport();
    }
}

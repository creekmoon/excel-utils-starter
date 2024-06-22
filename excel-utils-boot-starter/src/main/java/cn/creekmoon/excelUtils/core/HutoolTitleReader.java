package cn.creekmoon.excelUtils.core;

import cn.creekmoon.excelUtils.converter.StringConverter;
import cn.creekmoon.excelUtils.exception.CheckedExcelException;
import cn.creekmoon.excelUtils.exception.GlobalExceptionManager;
import cn.hutool.core.text.StrFormatter;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.sax.Excel07SaxReader;
import cn.hutool.poi.excel.sax.handler.RowHandler;
import jakarta.servlet.http.HttpServletResponse;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;

import java.io.IOException;
import java.util.*;
import java.util.function.BiConsumer;

import static cn.creekmoon.excelUtils.core.ExcelConstants.*;

@Slf4j
public class HutoolTitleReader<R> implements ITitleReader<R> {

    public SheetReaderContext sheetReaderContext;

    protected ExcelImport parent;

    protected SheetWriter sheetWriter;


    /**
     * 获取SHEET页的总行数
     *
     * @return
     */
    @SneakyThrows
    @Override
    public Long getSheetRowCount() {
        return getExcelImport().getSheetRowCount(getSheetReaderContext().sheetIndex);
    }

    @Override
    public <T> HutoolTitleReader<R> addConvert(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        getSheetReaderContext().title2converts.put(title, convert);
        getSheetReaderContext().title2consumers.put(title, setter);
        return this;
    }


    @Override
    public HutoolTitleReader<R> addConvert(String title, BiConsumer<R, String> reader) {
        addConvert(title, x -> x, reader);
        return this;
    }


    @Override
    public <T> HutoolTitleReader<R> addConvertAndSkipEmpty(String title, BiConsumer<R, String> setter) {
        getSheetReaderContext().skipEmptyTitles.add(title);
        addConvert(title, x -> x, setter);
        return this;
    }

    @Override
    public <T> HutoolTitleReader<R> addConvertAndSkipEmpty(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        getSheetReaderContext().skipEmptyTitles.add(title);
        addConvert(title, convert, setter);
        return this;
    }


    @Override
    public HutoolTitleReader<R> addConvertAndMustExist(String title, BiConsumer<R, String> setter) {
        getSheetReaderContext().mustExistTitles.add(title);
        addConvert(title, x -> x, setter);
        return this;
    }

    @Override
    public <T> HutoolTitleReader<R> addConvertAndMustExist(String title, ExFunction<String, T> convert, BiConsumer<R, T> setter) {
        getSheetReaderContext().mustExistTitles.add(title);
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
            this.getSheetReaderContext().convertPostProcessors.add(postProcessor);
        }
        return this;
    }


    @Override
    public ExcelImport read(ExConsumer<R> dataConsumer) {
        /* 以readAll模式进行读,再逐行消费*/
        List<R> rs = readAll();
        rs.forEach(x -> {
            try {
                dataConsumer.accept(x);
                this.setResult(x, IMPORT_SUCCESS_MSG);
            } catch (Exception e) {
                getExcelImport().getErrorCount().incrementAndGet();
                this.setResult(x, GlobalExceptionManager.getExceptionMsg(e));
            }
        });
        return getExcelImport();
    }

    @Override
    public TitleReadResult<R> read() throws InterruptedException {
        TitleReadResult<R> readResult = new TitleReadResult<>();
        parent.sheetIndex2ReadResult.put(getSheetReaderContext().sheetIndex, readResult);

        /*尝试拿锁*/
        ExcelImport.importSemaphore.acquire();
        try {
            //新版读取 使用SAX读取模式
            Excel07SaxReader excel07SaxReader = initSaxReader(getSheetReaderContext().sheetIndex, readResult);
            /*第一个参数 文件流  第二个参数 -1就是读取所有的sheet页*/
            excel07SaxReader.read(this.getExcelImport().file.getInputStream(), -1);
        } catch (Exception e) {
            log.error("SaxReader读取Excel文件异常", e);
            ExcelExport.cleanTempFileDelay(getExcelImport().excelExport.stopWrite());
        } finally {
            /*释放信号量*/
            ExcelImport.importSemaphore.release();
        }
        return readResult;
    }


//    @Deprecated
//    @SneakyThrows
//    public ExcelImport readByRow(ExConsumer<R> dataConsumer) {
//        /*尝试拿锁*/
//        ExcelImport.importSemaphore.acquire();
//        try {
//            //新版读取 使用SAX读取模式
//            Excel07SaxReader excel07SaxReader = initSaxReader(getSheetReaderContext().sheetIndex,
//                    (obj, rawMap) -> {
//                        dataConsumer.accept(obj);
//                        rawMap.put(RESULT_TITLE, IMPORT_SUCCESS_MSG);
//                    },
//                    (obj, rawMap) -> {
//                        sheetWriter.writeByMap(Collections.singletonList(rawMap));
//                        //getExcelImport().excelExport.getBigExcelWriter().write(Collections.singletonList(rawMap));
//                    });
//            sheetWriter = getExcelImport().excelExport.switchSheet(ExcelExport.generateSheetNameByIndex(getSheetReaderContext().sheetIndex), Map.class);
//            sheetWriter.setColumnWidthDefault();
//            /*第一个参数 文件流  第二个参数 -1就是读取所有的sheet页*/
//            excel07SaxReader.read(this.getExcelImport().file.getInputStream(), -1);
//        } catch (Exception e) {
//            log.error("SaxReader读取Excel文件异常", e);
//            e.printStackTrace();
//            ExcelExport.cleanTempFileDelay(getExcelImport().excelExport.stopWrite());
//        } finally {
//            /*释放信号量*/
//            ExcelImport.importSemaphore.release();
//        }
//        return this.getExcelImport();
//    }


//    /**
//     * 读取Excel内容到一个List (内存模式)
//     * 这个方法保证所有的数据都是通过格式校验的, 如果任一格式校验失败将返回整个空数组.
//     *
//     * @return
//     */
//    @Override
//    @SneakyThrows
//    public List<R> readAll() {
//        TitleReadResult<R> readResult = new TitleReadResult<>();
//        parent.sheetIndex2ReadResult.put(getSheetReaderContext().sheetIndex, readResult);
//
//        /*尝试拿锁*/
//        ExcelImport.importSemaphore.acquire();
//        try {
//            //新版读取 使用SAX读取模式
//            Excel07SaxReader excel07SaxReader = initSaxReader(getSheetReaderContext().sheetIndex, readResult);
//            /*第一个参数 文件流  第二个参数 -1就是读取所有的sheet页*/
//            excel07SaxReader.read(this.getExcelImport().file.getInputStream(), -1);
//        } catch (Exception e) {
//            log.error("SaxReader读取Excel文件异常", e);
//            ExcelExport.cleanTempFileDelay(getExcelImport().excelExport.stopWrite());
//        } finally {
//            /*释放信号量*/
//            ExcelImport.importSemaphore.release();
//        }
//
//        //假如存在convert阶段就失败的单, 说明readAll无法读取完整的数据, 此时将返回空集合
//        return readResult.getAll();
//    }


    /**
     * 增加读取范围限制
     *
     * @param titleRowIndex    标题所在的行数(下标按照从0开始, 如果是第一行则填0)
     * @param lastDataRowIndex 最后一条数据所在的行数(下标按照从0开始, 如果是第一行则填0)
     * @return
     */
    @Override
    public HutoolTitleReader<R> range(int titleRowIndex, int firstDataRowIndex, int lastDataRowIndex) {
        this.getSheetReaderContext().titleRowIndex = titleRowIndex;
        this.getSheetReaderContext().firstRowIndex = firstDataRowIndex;
        this.getSheetReaderContext().latestRowIndex = lastDataRowIndex;
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
        if (getSheetReaderContext().ENABLE_TITLE_CHECK) {
            if (getSheetReaderContext().TITLE_CHECK_FAIL_FLAG || !titleConsistencyCheck(getSheetReaderContext().title2converts.keySet(), row.keySet())) {
                getSheetReaderContext().TITLE_CHECK_FAIL_FLAG = true;
                throw new CheckedExcelException(TITLE_CHECK_ERROR);
            }
        }
        getSheetReaderContext().ENABLE_TITLE_CHECK = false;

        /*过滤空白行*/
        if (getSheetReaderContext().ENABLE_BLANK_ROW_FILTER
                && row.values().stream().allMatch(x -> x == null || "".equals(x))
        ) {
            return null;
        }

        /*初始化空对象*/
        R convertObject = (R) this.getSheetReaderContext().newObjectSupplier.get();
        /*最大转换次数*/
        int maxConvertCount = this.getSheetReaderContext().title2consumers.keySet().size();
        /*执行convert*/
        for (Map.Entry<String, Object> entry : row.entrySet()) {
            /*如果包含不支持的标题,  或者已经超过最大次数则不再进行读取*/
            if (!this.getSheetReaderContext().title2consumers.containsKey(entry.getKey()) || maxConvertCount-- <= 0) {
                continue;
            }
            String value = Optional.ofNullable(entry.getValue()).map(x -> (String) x).orElse("");
            /*检查必填项/检查可填项*/
            if (StrUtil.isBlank(value)) {
                if (this.getSheetReaderContext().mustExistTitles.contains(entry.getKey())) {
                    throw new CheckedExcelException(StrFormatter.format(FIELD_LACK_MSG, entry.getKey()));
                }
                if (this.getSheetReaderContext().skipEmptyTitles.contains(entry.getKey())) {
                    continue;
                }
            }
            /*转换数据*/
            try {
                Object convertValue = this.getSheetReaderContext().title2converts.get(entry.getKey()).apply(value);
                this.getSheetReaderContext().title2consumers.get(entry.getKey()).accept(convertObject, convertValue);
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
     * @param titleReadResult  读取的结果集
     * @return
     */
    Excel07SaxReader initSaxReader(int targetSheetIndex, TitleReadResult titleReadResult) {

        HashMap<Integer, String> colIndex2Title = new HashMap<>();
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
                if (rowIndex == getSheetReaderContext().titleRowIndex) {
                    for (int colIndex = 0; colIndex < rowList.size(); colIndex++) {
                        colIndex2Title.put(colIndex, StringConverter.parse(rowList.get(colIndex)));
                    }
                    return;
                }
                /*只读取指定范围的数据 */
                if (rowIndex == (int) getSheetReaderContext().titleRowIndex
                        || rowIndex < getSheetReaderContext().firstRowIndex
                        || rowIndex > getSheetReaderContext().latestRowIndex) {
                    return;
                }
                /*没有添加 convert直接跳过 */
                if (getSheetReaderContext().title2converts.isEmpty()
                        && getSheetReaderContext().title2consumers.isEmpty()
                ) {
                    return;
                }

                /*Excel解析原生的数据*/
                HashMap<String, Object> rowData = new LinkedHashMap<>();
                for (int colIndex = 0; colIndex < rowList.size(); colIndex++) {
                    rowData.put(colIndex2Title.get(colIndex), StringConverter.parse(rowList.get(colIndex)));
                }
                titleReadResult.rowIndex2rawData.put(rowIndex, rowData);
                /*转换成业务对象*/
                R currentObject = null;
                try {
                    /*转换*/
                    currentObject = rowConvert(rowData);
                    if (currentObject == null) {
                        return;
                    }
                    /*转换后置处理器*/
                    for (ExConsumer convertPostProcessor : getSheetReaderContext().convertPostProcessors) {
                        convertPostProcessor.accept(currentObject);
                    }
                    rowData.put(RESULT_TITLE, CONVERT_SUCCESS_MSG);
                    /*消费*/
                    titleReadResult.rowIndex2data.put(rowIndex, currentObject);
                } catch (Exception e) {
                    getExcelImport().getErrorCount().incrementAndGet();
                    titleReadResult.errorCount.incrementAndGet();
                    /*写入导出Excel结果*/
                    rowData.put(RESULT_TITLE, GlobalExceptionManager.getExceptionMsg(e));
                }
                if (currentObject == null && rowData != null) {
                    //假如存在任一数据convert阶段就失败的单, 将打一个标记
                    titleReadResult.EXISTS_CONVERT_FAIL.set(true);
                }
                getExcelImport().convertObject2rawData.put(currentObject, rowData);
                getExcelImport().sheetIndex2rawData
                        .computeIfAbsent(getSheetReaderContext().sheetIndex, k -> new ArrayList<>())
                        .add(rowData);

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
     * 设置读取的结果
     *
     * @param object 实例化的对象
     * @param msg    结果
     */
    public void setResult(R object, String msg) {
        getExcelImport().setResult(object, msg);
    }


    public ExcelImport response(HttpServletResponse response) throws IOException {
        return getExcelImport().response(response);
    }

    /**
     * 禁用标题一致性检查
     *
     * @return
     */
    @Override
    public HutoolTitleReader<R> disableTitleConsistencyCheck() {
        this.getSheetReaderContext().ENABLE_TITLE_CHECK = false;
        return this;
    }

    /**
     * 禁用空白行过滤
     *
     * @return
     */
    @Override
    public HutoolTitleReader<R> disableBlankRowFilter() {
        this.getSheetReaderContext().ENABLE_BLANK_ROW_FILTER = false;
        return this;
    }

    @Override
    public SheetReaderContext getSheetReaderContext() {
        return sheetReaderContext;
    }


    @Override
    public ExcelImport getExcelImport() {
        return getExcelImport();
    }
}

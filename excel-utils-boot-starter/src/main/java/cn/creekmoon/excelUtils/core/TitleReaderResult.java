package cn.creekmoon.excelUtils.core;

import cn.creekmoon.excelUtils.exception.GlobalExceptionManager;
import lombok.extern.slf4j.Slf4j;
import org.jetbrains.annotations.Nullable;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;


/**
 * 读取结果
 */
@Slf4j
public class TitleReaderResult<R> extends ReaderResult {


    AtomicReference<Boolean> EXISTS_CONVERT_FAIL = new AtomicReference<>(false);

    /*行对象集合*/
    LinkedHashMap<Integer, R> rowIndex2data = new LinkedHashMap<>();

    /*行结果集合*/
    LinkedHashMap<Integer, String> rowIndex2msg = new LinkedHashMap<>();

    /*行对象集合*/
    LinkedHashMap<R, Integer> data2rowIndex = new LinkedHashMap<>();

    /*原始行对象集合*/
    LinkedHashMap<Integer, Map<String, Object>> rowIndex2rawData = new LinkedHashMap<>();


    public List<R> getAll() {
        if (EXISTS_CONVERT_FAIL.get()) {
            // 如果转化阶段就存在失败数据, 意味着数据不完整,应该返回空
            return new ArrayList<>();
        }
        return new ArrayList<>(rowIndex2data.values());
    }


    public TitleReaderResult<R> foreachAndConsume(ExConsumer<R> dataConsumer) {
        return this.foreachAndConsume((index, data) -> dataConsumer.accept(data));
    }

    public TitleReaderResult<R> foreachAndConsume(ExBiConsumer<Integer, R> rowIndexAndDataConsumer) {
        rowIndex2data.forEach((rowIndex, data) -> {
            try {
                rowIndexAndDataConsumer.accept(rowIndex, data);
                rowIndex2msg.put(rowIndex, ExcelConstants.IMPORT_SUCCESS_MSG);
            } catch (Exception e) {
                String exceptionMsg = GlobalExceptionManager.getExceptionMsg(e);
                rowIndex2msg.put(rowIndex, exceptionMsg);
            }
        });
        return this;
    }


    public TitleReaderResult<R> setResultMsg(R data, String msg) {
        Integer i = getDataIndexOrNull(data);
        if (i == null) {
            return null;
        }
        rowIndex2msg.put(i, msg);
        return this;
    }


    private @Nullable Integer getDataIndexOrNull(R data) {
        if (data2rowIndex.size() != rowIndex2data.size()) {
            log.error("[Excel读取异常]请确保实体类{}不能覆写Equal和HashCode注解.", data.getClass().toString());
        }
        Integer i = data2rowIndex.get(data);
        if (i == null) {
            log.error("[Excel读取异常]对象不在读取结果中! [{}]", data);
            return null;
        }
        return i;
    }


    public String getResultMsg(R data) {
        Integer i = getDataIndexOrNull(data);
        if (i == null) return null;
        return rowIndex2msg.get(i);
    }


    public TitleReaderResult<R> setResultMsg(Integer rowIndex, String msg) {
        if (rowIndex2msg.containsKey(rowIndex)) {
            rowIndex2msg.put(rowIndex, msg);
        }
        return this;
    }

    public String getResultMsg(Long rowIndex) {
        return rowIndex2msg.get(Math.toIntExact(rowIndex));
    }

    public Integer getErrorCount() {
        return errorCount.get();
    }

    public Integer getDataLatestRowIndex() {
        return this.rowIndex2msg.keySet().stream().max(Integer::compareTo).get();
    }

    public Integer getDataFirstRowIndex() {
        return this.rowIndex2msg.keySet().stream().min(Integer::compareTo).get();
    }
}

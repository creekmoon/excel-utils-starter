package cn.creekmoon.excelUtils.core;

import cn.creekmoon.excelUtils.exception.GlobalExceptionManager;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;


/**
 * 读取结果
 */
public class TitleReaderResult<R> extends ReaderResult {


    AtomicReference<Boolean> EXISTS_CONVERT_FAIL = new AtomicReference<>(false);
    /*行对象集合*/
    LinkedHashMap<Integer, R> rowIndex2data = new LinkedHashMap<>();

    /*原始行对象集合*/
    LinkedHashMap<Integer, Map<String, Object>> rowIndex2rawData = new LinkedHashMap<>();
    /*原始行对象集合*/
    LinkedHashMap<Integer, String> rowIndex2msg = new LinkedHashMap<>();

    /*原始行对象集合*/
    LinkedHashMap<R, Map<String, Object>> data2RawData = new LinkedHashMap<>();


    public List<R> getAll() {
        if (EXISTS_CONVERT_FAIL.get()) {
            // 如果转化阶段就存在失败数据, 意味着数据不完整,应该返回空
            return new ArrayList<>();
        }
        return new ArrayList<>(rowIndex2data.values());
    }


    public TitleReaderResult<R> foreach(ExConsumer<R> dataConsumer) {
        return this.foreach((index, data) -> dataConsumer.accept(data));
    }

    public TitleReaderResult<R> foreach(ExBiConsumer<Integer, R> rowIndexAndDataConsumer) {
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
        return null;
    }

    public String getResultMsg(R data) {
        return "";
    }


    public TitleReaderResult<R> setResultMsg(Integer rowIndex, String msg) {
        return null;
    }

    public String getResultMsg(Long rowIndex) {
        return "";
    }

    public Long getErrorCount() {
        return 0L;
    }

    public Integer getDataLatestRowIndex() {
        return this.rowIndex2msg.keySet().stream().max(Integer::compareTo).get();
    }

    public Integer getDataFirstRowIndex() {
        return this.rowIndex2msg.keySet().stream().min(Integer::compareTo).get();
    }
}

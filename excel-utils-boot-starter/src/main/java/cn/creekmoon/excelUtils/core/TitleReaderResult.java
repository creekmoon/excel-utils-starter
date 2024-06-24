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
public class TitleReaderResult<R> extends ReaderResult implements ITitleReaderResult<R> {


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


    @Override
    public ITitleReaderResult<R> foreach(ExConsumer<R> dataConsumer) {
        return this.foreach((index, data) -> dataConsumer.accept(data));
    }

    @Override
    public ITitleReaderResult<R> foreach(ExBiConsumer<Integer, R> rowIndexAndDataConsumer) {
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

    @Override
    public ITitleReaderResult<R> setResultMsg(R data, String msg) {
        return null;
    }

    @Override
    public String getResultMsg(R data) {
        return "";
    }


    @Override
    public ITitleReaderResult<R> setResultMsg(Integer rowIndex, String msg) {
        return null;
    }

    @Override
    public String getResultMsg(Long rowIndex) {
        return "";
    }

    @Override
    public Long getErrorCount() {
        return 0L;
    }

    @Override
    public Integer getDataLatestRowIndex() {
        return this.rowIndex2msg.keySet().stream().max(Integer::compareTo).get();
    }

    @Override
    public Integer getDataFirstRowIndex() {
        return this.rowIndex2msg.keySet().stream().min(Integer::compareTo).get();
    }
}

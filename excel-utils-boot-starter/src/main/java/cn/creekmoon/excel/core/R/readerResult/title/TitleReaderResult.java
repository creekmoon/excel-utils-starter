package cn.creekmoon.excel.core.R.readerResult.title;

import cn.creekmoon.excel.core.R.readerResult.ReaderResult;
import cn.creekmoon.excel.util.ExcelConstants;
import cn.creekmoon.excel.util.exception.ExBiConsumer;
import cn.creekmoon.excel.util.exception.ExConsumer;
import cn.creekmoon.excel.util.exception.GlobalExceptionMsgManager;
import cn.hutool.core.map.BiMap;
import cn.hutool.core.text.StrFormatter;
import lombok.extern.slf4j.Slf4j;
import org.jetbrains.annotations.Nullable;

import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;


/**
 * 读取结果
 */
@Slf4j
public class TitleReaderResult<R> implements ReaderResult<R> {

    public LocalDateTime readStartTime;
    public LocalDateTime readSuccessTime;
    public LocalDateTime consumeSuccessTime;

    /*错误次数统计*/
    public AtomicInteger errorCount = new AtomicInteger(0);
    public StringBuilder errorReport = new StringBuilder();

    /*存在读取失败的数据*/
    public AtomicReference<Boolean> EXISTS_READ_FAIL = new AtomicReference<>(false);

    /*K=行下标 V=数据*/
    public BiMap<Integer, R> rowIndex2dataBiMap = new BiMap<>(new LinkedHashMap<>());

    /*行结果集合*/
    public LinkedHashMap<Integer, String> rowIndex2msg = new LinkedHashMap<>();

    public List<R> getAll() {
        if (EXISTS_READ_FAIL.get()) {
            // 如果转化阶段就存在失败数据, 意味着数据不完整,应该返回空
            return new ArrayList<>();
        }
        return new ArrayList<>(rowIndex2dataBiMap.values());
    }


    public TitleReaderResult<R> consume(ExConsumer<R> dataConsumer) {
        return this.consume((index, data) -> dataConsumer.accept(data));
    }

    public TitleReaderResult<R> consume(ExBiConsumer<Integer, R> rowIndexAndDataConsumer) {
        rowIndex2dataBiMap.forEach((rowIndex, data) -> {
            try {
                rowIndexAndDataConsumer.accept(rowIndex, data);
                rowIndex2msg.put(rowIndex, ExcelConstants.IMPORT_SUCCESS_MSG);
            } catch (Exception e) {
                errorCount.incrementAndGet();
                String exceptionMsg = GlobalExceptionMsgManager.getExceptionMsg(e);
                getErrorReport().append(StrFormatter.format("第[{}]行发生错误[{}]", (int) rowIndex + 1, exceptionMsg));
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

        Integer i = rowIndex2dataBiMap.getKey(data);
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


    public String getResultMsg(Integer rowIndex) {
        return rowIndex2msg.get(rowIndex);
    }

    public TitleReaderResult<R> setResultMsg(Integer rowIndex, String msg) {
        if (rowIndex2msg.containsKey(rowIndex)) {
            rowIndex2msg.put(rowIndex, msg);
        }
        return this;
    }


    @Override
    public StringBuilder getErrorReport() {
        return errorReport;
    }

    public AtomicInteger getErrorCount() {
        return errorCount;
    }

    @Override
    public LocalDateTime getReadStartTime() {
        return readStartTime;
    }

    @Override
    public LocalDateTime getReadSuccessTime() {
        return readSuccessTime;
    }

    @Override
    public LocalDateTime getConsumeSuccessTime() {
        return consumeSuccessTime;
    }


    public Integer getDataLatestRowIndex() {
        return this.rowIndex2msg.keySet().stream().max(Integer::compareTo).get();
    }

    public Integer getDataFirstRowIndex() {
        return this.rowIndex2msg.keySet().stream().min(Integer::compareTo).get();
    }
}

package cn.creekmoon.excel.core.R.readerResult.cell;

import cn.creekmoon.excel.core.R.readerResult.ReaderResult;
import cn.creekmoon.excel.util.exception.ExConsumer;

import java.time.LocalDateTime;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

public class CellReaderResult<R> implements ReaderResult<R> {

    public LocalDateTime readStartTime;
    public LocalDateTime readSuccessTime;
    public LocalDateTime consumeSuccessTime;
    public AtomicReference<Boolean> EXISTS_READ_FAIL = new AtomicReference<>(false);
    /*错误次数统计*/
    public AtomicInteger errorCount = new AtomicInteger(0);
    public StringBuilder errorReport = new StringBuilder();

    public R data = null;

    @Override
    public StringBuilder getErrorReport() {
        return errorReport;
    }

    @Override
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

    @Override
    public ReaderResult<R> consume(ExConsumer<R> consumer) throws Exception {
        if (data == null) {
            return this;
        }
        consumer.accept(data);
        consumeSuccessTime = LocalDateTime.now();
        return this;
    }


    public R getData() {
        if (EXISTS_READ_FAIL.get()) {
            return null;
        }
        return data;
    }

}

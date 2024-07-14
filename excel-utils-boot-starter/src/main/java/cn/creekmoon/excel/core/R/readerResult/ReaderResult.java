package cn.creekmoon.excel.core.R.readerResult;


import cn.creekmoon.excel.util.exception.ExConsumer;

import java.time.LocalDateTime;
import java.util.concurrent.atomic.AtomicInteger;

/**
 * 读取结果
 */
public interface ReaderResult<R> {

    public StringBuilder getErrorReport();

    public AtomicInteger getErrorCount();

    public LocalDateTime getReadStartTime();

    public LocalDateTime getReadSuccessTime();

    public LocalDateTime getConsumeSuccessTime();

    ReaderResult<R> consume(ExConsumer<R> consumer) throws Exception;
}

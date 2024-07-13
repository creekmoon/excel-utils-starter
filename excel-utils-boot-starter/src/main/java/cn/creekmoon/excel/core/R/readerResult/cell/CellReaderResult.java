package cn.creekmoon.excel.core.R.readerResult.cell;

import cn.creekmoon.excel.core.R.readerResult.ReaderResult;

import java.util.concurrent.atomic.AtomicInteger;

public class CellReaderResult<R> implements ReaderResult {

    /*错误次数统计*/
    public AtomicInteger errorCount = new AtomicInteger(0);

    public Integer durationSecond = 0;


    @Override
    public Integer getErrorCount() {
        return errorCount.get();
    }

    @Override
    public Integer getDurationSecond() {
        return 0;
    }
}

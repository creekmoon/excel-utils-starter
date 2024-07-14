package cn.creekmoon.excel.core.R.reader;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.readerResult.ReaderResult;
import cn.creekmoon.excel.util.exception.ExConsumer;

import java.io.IOException;

public interface Reader<R> {


    ReaderContext getReaderContext();

    Integer getSheetIndex();

    ExcelImport getExcelImport();

    ReaderResult getReadResult();

    ReaderResult<R> read(ExConsumer<R> consumer) throws Exception;

    ReaderResult<R> read() throws InterruptedException, IOException;

}

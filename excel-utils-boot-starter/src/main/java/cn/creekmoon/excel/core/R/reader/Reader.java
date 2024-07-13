package cn.creekmoon.excel.core.R.reader;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.readerResult.ReaderResult;

public interface Reader<R> {


    ReaderContext getReaderContext();

    Integer getSheetIndex();

    ExcelImport getExcelImport();

    ReaderResult getReadResult();

}

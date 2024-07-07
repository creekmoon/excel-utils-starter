package cn.creekmoon.excelUtils.core.reader;

import cn.creekmoon.excelUtils.core.ExcelImport;
import cn.creekmoon.excelUtils.core.ReaderContext;
import cn.creekmoon.excelUtils.core.ReaderResult;

public interface Reader<R> {


    ReaderContext getReaderContext();

    Integer getSheetIndex();

    ExcelImport getExcelImport();

    ReaderResult getReadResult();

}

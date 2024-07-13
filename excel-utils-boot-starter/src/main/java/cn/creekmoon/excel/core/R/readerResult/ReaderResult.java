package cn.creekmoon.excel.core.R.readerResult;


/**
 * 读取结果
 */
public interface ReaderResult<R> {


    public Integer getErrorCount();

    //持续时长 秒
    public Integer getDurationSecond();


}

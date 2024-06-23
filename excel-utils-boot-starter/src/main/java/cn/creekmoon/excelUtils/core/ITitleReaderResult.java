package cn.creekmoon.excelUtils.core;

public interface ITitleReaderResult<R> {

    /**
     * 消费数据 和 设置结果
     */
    ITitleReaderResult<R> foreach(ExConsumer<R> dataConsumer);

    ITitleReaderResult<R> setResultMsg(R data, String msg);

    String getResultMsg(R data);


    /**
     * 消费数据(以下标形式) 和 设置结果
     */
    ITitleReaderResult<R> foreach(ExBiConsumer<Integer, R> rowIndexAndDataConsumer);

    ITitleReaderResult<R> setResultMsg(Integer rowIndex, String msg);

    String getResultMsg(Long rowIndex);


    /**
     * 获取错误数量
     *
     * @return
     */
    Long getErrorCount();

}

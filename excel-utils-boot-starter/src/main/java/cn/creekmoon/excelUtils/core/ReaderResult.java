package cn.creekmoon.excelUtils.core;

import java.time.LocalDateTime;
import java.util.concurrent.atomic.AtomicInteger;


/**
 * 读取结果
 *
 *
 * */
public class ReaderResult<R> {


    /*读取时间审计*/
    protected LocalDateTime readStartTime = LocalDateTime.now();
    protected LocalDateTime readEndTime = LocalDateTime.now();

    /*对应的Sheet页标签*/
    protected Integer sheerIndex;

    /*错误次数统计*/
    protected AtomicInteger errorCount = new AtomicInteger(0);

    /*单元格错误信息统计*/
//    protected LinkedHashMap<Integer,LinkedHashMap<Integer,String>> rowIndex2ColIndex2ErrorMsg = new LinkedHashMap<>();


}

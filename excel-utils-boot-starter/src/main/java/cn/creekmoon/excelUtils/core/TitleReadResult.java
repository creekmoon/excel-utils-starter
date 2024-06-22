package cn.creekmoon.excelUtils.core;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.atomic.AtomicReference;


/**
 * 读取结果
 */
public class TitleReadResult<R> extends ReadResult {


    AtomicReference<Boolean> EXISTS_CONVERT_FAIL = new AtomicReference<>(false);
    /*行对象集合*/
    LinkedHashMap<Integer, R> rowIndex2data = new LinkedHashMap<>();

    /*原始行对象集合*/
    LinkedHashMap<Integer, Map<String, Object>> rowIndex2rawData = new LinkedHashMap<>();
    /*原始行对象集合*/
    LinkedHashMap<R, Map<String, Object>> data2RawData = new LinkedHashMap<>();


    List<R> getAll() {
        if (EXISTS_CONVERT_FAIL.get()) {
            // 如果转化阶段就存在失败数据, 意味着数据不完整,应该返回空
            return new ArrayList<>();
        }
        return new ArrayList<>(rowIndex2data.values());
    }


}

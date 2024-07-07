package cn.creekmoon.excelUtils.core;


import cn.hutool.core.lang.UUID;
import cn.hutool.poi.excel.BigExcelWriter;
import cn.hutool.poi.excel.ExcelUtil;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.jetbrains.annotations.NotNull;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

/**
 * @author JY
 * @date 2022-01-05
 */
@Slf4j
public class ExcelExport {
    /**
     * 打印调试内容
     */
    protected boolean debugger = false;
    /*唯一识别名称*/
    public String taskId = UUID.fastUUID().toString();
    /*自定义的名称*/
    public String excelName;
    /*写入器*/
    private BigExcelWriter bigExcelWriter;
    /*
     * sheet页和导出对象的映射关系
     * */
    private Map<String, SheetWriter> sheetName2SheetWriter = new HashMap<>();

    private ExcelExport() {

    }

    public static <T> SheetWriter<T> create(Class<T> c) {
        ExcelExport excelExport = create();
        return excelExport.switchSheet(c);
    }

    public static ExcelExport create() {
        ExcelExport excelExport = new ExcelExport();
        excelExport.excelName = ExcelConstants.excelNameGenerator.get();
        return excelExport;
    }

    /**
     * 生成sheetName
     *
     * @param sheetIndex sheet页下标
     * @return
     */
    @NotNull
    protected static String generateSheetNameByIndex(Integer sheetIndex) {
        return "Sheet" + (sheetIndex + 1);
    }

    /**
     * 切换到新的标签页
     */
    public <T> SheetWriter<T> switchSheet(String sheetName, Class<T> newDataClass) {
        /*第一次切换sheet页是重命名当前sheet*/
        if (sheetName2SheetWriter.isEmpty()) {
            getBigExcelWriter().renameSheet(sheetName);
        }
        /*后续切换sheet页是新增sheet*/
        if (!sheetName2SheetWriter.isEmpty()) {
            getBigExcelWriter().setSheet(sheetName);
        }
        return sheetName2SheetWriter.computeIfAbsent(sheetName, s -> {
            return new SheetWriter<>(this, new SheetWriterContext(sheetName));
        });

    }

    /**
     * 切换到新的标签页
     */
    public <T> SheetWriter<T> switchSheet(Class<T> newDataClass) {
        int indexSeq = 0;
        while (sheetName2SheetWriter.containsKey(generateSheetNameByIndex(indexSeq))) {
            indexSeq++;
        }
        return switchSheet(generateSheetNameByIndex(indexSeq), newDataClass);
    }

    public ExcelExport debug() {
        this.debugger = true;
        return this;
    }


    /**
     * 停止写入
     *
     * @return taskId
     */
    public String stopWrite() {
        getBigExcelWriter().close();
        return taskId;
    }







    /**
     * 响应并清除文件
     *
     * @param response
     * @throws IOException
     */
    public void response(HttpServletResponse response) throws IOException {
        String taskId = this.stopWrite();
        ExcelFileUtils.response(ExcelFileUtils.getAbsoluteFilePath(taskId), excelName, response);
    }


    /**
     * 内部操作类,但是暴露出来了,希望最好不要用这个方法
     *
     * @return
     */
    public BigExcelWriter getBigExcelWriter() {
        if (bigExcelWriter == null) {
            bigExcelWriter = ExcelUtil.getBigWriter(ExcelFileUtils.getAbsoluteFilePath(taskId));
        }
        return bigExcelWriter;
    }

    /**
     * 写入策略
     */
    public enum WriteStrategy {
        /*忽略取值异常 通常是多级属性空指针造成的 如果取不到值直接置为NULL*/
        CONTINUE_ON_ERROR,
        /*遇到任何失败的情况则停止*/
        STOP_ON_ERROR;
    }

}

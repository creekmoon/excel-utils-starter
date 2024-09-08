package cn.creekmoon.excel.core.W;


import cn.creekmoon.excel.core.W.title.HutoolTitleWriter;
import cn.creekmoon.excel.util.ExcelFileUtils;
import cn.hutool.core.lang.UUID;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.jetbrains.annotations.NotNull;

import java.io.IOException;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
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
    public boolean debugger = false;
    /*唯一识别名称*/
    public String taskId = UUID.fastUUID().toString();
    /*生成的临时文件路径*/
    public String filePath = ExcelFileUtils.generateXlsxAbsoluteFilePath(taskId);
    /*自定义的名称*/
    public String excelName;


    /*
     * sheet页和导出对象的映射关系
     * */
    public Map<String, Writer> sheetName2SheetWriter = new HashMap<>();


    private ExcelExport() {

    }

    public static <T> HutoolTitleWriter<T> create(Class<T> c) {
        ExcelExport excelExport = create();
        return excelExport.switchSheet(c);
    }

    public static ExcelExport create() {
        ExcelExport excelExport = new ExcelExport();
        excelExport.excelName = "export_result_" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy_MM_dd_HH_mm"));
        return excelExport;
    }

    /**
     * 生成sheetName
     *
     * @param sheetIndex sheet页下标
     * @return
     */
    @NotNull
    protected static String getDefaultIndexName(Integer sheetIndex) {
        return "Sheet" + (sheetIndex + 1);
    }

    /**
     * 切换到新的标签页
     */
    public <T> HutoolTitleWriter<T> switchSheet(String sheetName, Class<T> newDataClass) {
        /*第一次切换sheet页是重命名当前sheet*/
        if (sheetName2SheetWriter.isEmpty()) {
            getBigExcelWriter().renameSheet(sheetName);
        }
        /*后续切换sheet页是新增sheet*/
        if (!sheetName2SheetWriter.isEmpty()) {
            getBigExcelWriter().setSheet(sheetName);
        }
        return sheetName2SheetWriter.computeIfAbsent(sheetName, s -> {
            return new HutoolTitleWriter<T>(this, sheetName);
        });

    }


    /**
     * 切换到新的标签页
     */
    public <T> HutoolTitleWriter<T> switchSheet(Class<T> newDataClass) {
        int indexSeq = 0;
        while (sheetName2SheetWriter.containsKey(getDefaultIndexName(indexSeq))) {
            indexSeq++;
        }
        return switchSheet(getDefaultIndexName(indexSeq), newDataClass);
    }

    public ExcelExport debug() {
        this.debugger = true;
        return this;
    }


    /**
     * 响应并清除文件
     *
     * @param response
     * @throws IOException
     */
    public void response(HttpServletResponse response) throws IOException {
        String taskId = this.stopWrite();
        ExcelFileUtils.response(ExcelFileUtils.generateXlsxAbsoluteFilePath(taskId), excelName, response);
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

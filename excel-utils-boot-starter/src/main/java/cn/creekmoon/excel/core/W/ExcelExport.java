package cn.creekmoon.excel.core.W;


import cn.creekmoon.excel.core.W.title.HutoolTitleWriter;
import cn.creekmoon.excel.core.W.title.TitleWriter;
import cn.creekmoon.excel.util.ExcelFileUtils;
import cn.hutool.core.lang.UUID;
import cn.hutool.core.map.BiMap;
import cn.hutool.core.text.StrFormatter;
import jakarta.servlet.http.HttpServletResponse;
import lombok.Getter;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.CellStyle;

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
    @Getter
    protected String taskId = UUID.fastUUID().toString();
    /*生成的临时文件路径*/
    @Getter
    protected String resultFilePath = ExcelFileUtils.generateXlsxAbsoluteFilePath(taskId);

    /*自定义的名称*/
    public String excelName;


    /* key=内置的样式风格   value=当前运行时对象里生成的实际样式
     *
     *   这里由各个 writer进行动态维护,他们应该在合适的时机将样式加进来
     **/
    public Map<PresetCellStyle, CellStyle> cellStyle2RunningTimeStyleObject;


    /*
     * sheet页和导出对象的映射关系
     * */
    public BiMap<Integer, Writer> sheetIndex2SheetWriter = new BiMap<>(new HashMap<>());


    private ExcelExport() {

    }


    public static ExcelExport create() {
        ExcelExport excelExport = new ExcelExport();
        excelExport.excelName = "export_result_" + LocalDateTime.now().format(DateTimeFormatter.ofPattern("yyyy_MM_dd_HH_mm"));
        return excelExport;
    }


    /**
     * 切换到新的标签页
     */
    public <T> TitleWriter<T> switchSheet(Integer sheetIndex, Class<T> newDataClass) {
        if (sheetIndex2SheetWriter.containsKey(sheetIndex)) {
            return (TitleWriter<T>) sheetIndex2SheetWriter.get(sheetIndex);
        }
        if (!sheetIndex2SheetWriter.containsKey(sheetIndex - 1)) {
            throw new RuntimeException(StrFormatter.format("切换sheet页失败,预期切换到sheet索引={}, 但当前最大索引下标sheet={},请检查代码并保证sheet页下标连续!"));
        }
        return switchNewSheet(newDataClass);
    }


    /**
     * 切换到新的标签页
     */
    public <T> TitleWriter<T> switchNewSheet(Class<T> newDataClass) {
        HutoolTitleWriter<T> newTitleWriter = new HutoolTitleWriter<T>(this, sheetIndex2SheetWriter.size());
        sheetIndex2SheetWriter.put(newTitleWriter.getSheetIndex(), newTitleWriter);
        return newTitleWriter;
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
        ExcelFileUtils.response(this.stopWrite(), excelName, response);
    }


    /**
     * 停止写入
     *
     * @return 结果文件绝对路径
     */
    public String stopWrite() {
        sheetIndex2SheetWriter.values().forEach(Writer::unsafeOnStopWrite);
        return getResultFilePath();
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

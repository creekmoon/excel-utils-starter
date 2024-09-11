package cn.creekmoon.excel.core.W.title.ext;

import cn.creekmoon.excel.core.W.ExcelExport;

import java.util.function.Predicate;

/*条件样式*/

public class ConditionStyle<R> {
    public Predicate<R> condition;
    public ExcelExport.PresetCellStyle presetCellStyle;
}

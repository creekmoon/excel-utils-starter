package cn.creekmoon.excel.core.W.title.ext;

import cn.creekmoon.excel.core.W.ExcelExport;
import lombok.AllArgsConstructor;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.function.Predicate;

/*条件样式*/

public class ConditionStyle<R> {
    public Predicate<R> condition;
    public ExcelExport.PresetCellStyle presetCellStyle;
    public CellStyle runningTimeStyle;

    //todo 思考如何平衡 PresetCellStyle 和  runningTimeStyle 的关系, 是在addTitle这个地方完成初始化吗?

}

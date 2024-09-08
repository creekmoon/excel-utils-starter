package cn.creekmoon.excel.core.W.title.ext;

import lombok.AllArgsConstructor;
import org.apache.poi.ss.usermodel.CellStyle;

import java.util.function.Predicate;

/*条件样式*/
@AllArgsConstructor
public class ConditionStyle<R> {
    Predicate<R> condition;
    CellStyle style;
}

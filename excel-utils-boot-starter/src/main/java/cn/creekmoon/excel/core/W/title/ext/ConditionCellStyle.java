package cn.creekmoon.excel.core.W.title.ext;

import lombok.Getter;

import java.util.function.Predicate;

/*条件样式*/
@Getter
public class ConditionCellStyle<R> {
    protected Predicate<R> condition;
    protected DefaultCellStyle defaultCellStyle;

    ConditionCellStyle(DefaultCellStyle defaultCellStyle, Predicate<R> condition){
        this.defaultCellStyle = defaultCellStyle;
        this.condition = condition;
    }
}
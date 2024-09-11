package cn.creekmoon.excel.core.W;

import cn.creekmoon.excel.core.W.title.ext.ConditionCellStyle;
import cn.creekmoon.excel.core.W.title.ext.DefaultCellStyle;
import lombok.Getter;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

public abstract class Writer {

    @Getter
    protected Integer sheetIndex;


    abstract protected Workbook getWorkbook();


    /**
     * 获取运行时的样式对象
     */
    abstract protected CellStyle getRunningTimeCellStyle(DefaultCellStyle style);


    /**
     * 声明周期钩子函数, 当写入数据时
     */
    protected void unsafeOnWrite() {
    }

    ;

    /**
     * 声明周期钩子函数, 当切换sheet页时
     */
    protected void unsafeOnSwitchSheet() {
    }

    ;

    /**
     * 声明周期钩子函数, 当切换sheet页时
     */
    protected void unsafeOnStopWrite() {
    }

    ;
}

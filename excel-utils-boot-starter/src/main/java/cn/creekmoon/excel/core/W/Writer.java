package cn.creekmoon.excel.core.W;

import lombok.Getter;

public abstract class Writer {

    @Getter
    protected Integer sheetIndex;

    /**
     * 声明周期钩子函数, 当切换sheet页时
     */
    abstract protected void unsafeOnSwitchSheet();

    /**
     * 声明周期钩子函数, 当切换sheet页时
     */
    abstract protected void unsafeOnStopWrite();
}

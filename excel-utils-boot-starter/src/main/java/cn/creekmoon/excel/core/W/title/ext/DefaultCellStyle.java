package cn.creekmoon.excel.core.W.title.ext;

import lombok.Getter;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.function.Consumer;

public  class DefaultCellStyle {
    public final static DefaultCellStyle LIGHT_ORANGE = new DefaultCellStyle(cellStyle ->
    {
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });

    public final static DefaultCellStyle ROYAL_BLUE = new DefaultCellStyle(cellStyle ->
    {
        cellStyle.setFillForegroundColor(IndexedColors.ROYAL_BLUE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });

    @Getter
    private Consumer<CellStyle> styleInitializer;

    private DefaultCellStyle(Consumer<CellStyle> styleInitializer) {
        this.styleInitializer = styleInitializer;
    }


}
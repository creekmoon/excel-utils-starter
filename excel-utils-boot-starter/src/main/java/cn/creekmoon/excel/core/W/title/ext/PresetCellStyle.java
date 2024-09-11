package cn.creekmoon.excel.core.W.title.ext;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.function.Consumer;

public enum PresetCellStyle {
    LIGHT_ORANGE(cellStyle ->
    {
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });


    private Consumer<CellStyle> styleInitializer;

    private PresetCellStyle(Consumer<CellStyle> styleInitializer) {
        this.styleInitializer = styleInitializer;
    }


}
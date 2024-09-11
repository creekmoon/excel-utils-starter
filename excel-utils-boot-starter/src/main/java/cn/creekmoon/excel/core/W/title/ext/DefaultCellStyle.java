package cn.creekmoon.excel.core.W.title.ext;

import lombok.Getter;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;

import java.util.function.Consumer;

public enum DefaultCellStyle {
    LIGHT_ORANGE(cellStyle ->
    {
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });

    @Getter
    private Consumer<CellStyle> styleInitializer;

    private DefaultCellStyle(Consumer<CellStyle> styleInitializer) {
        this.styleInitializer = styleInitializer;
    }


}
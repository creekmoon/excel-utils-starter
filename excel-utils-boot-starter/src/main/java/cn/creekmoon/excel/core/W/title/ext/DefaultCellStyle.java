package cn.creekmoon.excel.core.W.title.ext;

import lombok.Getter;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.function.BiConsumer;
import java.util.function.Consumer;

public class DefaultCellStyle {

    public final static DefaultCellStyle LIGHT_ORANGE = new DefaultCellStyle(cellStyle ->
    {
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });


    public final static DefaultCellStyle PALE_BLUE = new DefaultCellStyle(cellStyle ->
    {
        cellStyle.setFillForegroundColor(IndexedColors.PALE_BLUE.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });


    public final static DefaultCellStyle LIGHT_GREEN = new DefaultCellStyle(cellStyle ->
    {
        cellStyle.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    });


    @Getter
    private BiConsumer<Workbook, CellStyle> styleInitializer;

    public DefaultCellStyle(Consumer<CellStyle> styleInitializer) {
        this.styleInitializer = (x, y) -> styleInitializer.accept(y);
    }


    public DefaultCellStyle(BiConsumer<Workbook, CellStyle> styleInitializer) {
        this.styleInitializer = styleInitializer;
    }
}
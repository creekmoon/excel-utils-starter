package cn.creekmoon.excelUtils.hutool589.poi.excel.cell.values;

import cn.creekmoon.excelUtils.hutool589.core.util.StrUtil;
import cn.creekmoon.excelUtils.hutool589.poi.excel.cell.CellValue;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.FormulaError;

/**
 * ERROR类型单元格值
 *
 * @author looly
 * @since 5.7.8
 */
public class ErrorCellValue implements CellValue<String> {

    private final Cell cell;

    /**
     * 构造
     *
     * @param cell {@link Cell}
     */
    public ErrorCellValue(Cell cell) {
        this.cell = cell;
    }

    @Override
    public String getValue() {
        final FormulaError error = FormulaError.forInt(cell.getErrorCellValue());
        return (null == error) ? StrUtil.EMPTY : error.getString();
    }
}
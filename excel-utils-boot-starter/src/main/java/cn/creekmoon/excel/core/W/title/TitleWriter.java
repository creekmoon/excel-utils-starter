package cn.creekmoon.excel.core.W.title;

import cn.creekmoon.excel.core.W.Writer;
import cn.creekmoon.excel.core.W.title.ext.ConditionStyle;
import cn.creekmoon.excel.core.W.title.ext.Title;
import cn.hutool.poi.excel.style.StyleUtil;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.function.Function;

public abstract class TitleWriter<R> extends Writer {

    /**
     * 表头集合
     */
    protected List<Title> titles = new ArrayList<>();

    /**
     * 表头集合
     */
    protected HashMap<Integer, List<ConditionStyle>> colIndex2Styles = new HashMap<>();

    /* 多级表头时会用到 全局标题深度  initTitle方法会给其赋值
     *
     *   |       title               |     深度=3    rowIndex=0
     *   |   titleA    |    titleB   |     深度=2    rowIndex=1
     *   |title1|title2|title3|title4|     深度=1    rowIndex=2
     * */
    protected Integer MAX_TITLE_DEPTH = null;

    /* 多级表头时会用到  深度和标题的映射关系*/
    protected HashMap<Integer, List<Title>> depth2Titles = new HashMap<>();


    /**
     * 添加标题
     */
    public TitleWriter<R> addTitle(String titleName, Function<R, Object> valueFunction) {
        titles.add(Title.of(titleName, valueFunction));
        return this;
    }

    public abstract int countTitles();


    public abstract HutoolTitleWriter<R> write(List<R> data);

    protected void unsafeInit() {
        Workbook workbook = getWorkbook();
        /*初始化样式*/
//        CellStyle newCellStyle = getBigExcelWriter().createCellStyle();
        XSSFCellStyle newCellStyle = (XSSFCellStyle) StyleUtil.createDefaultCellStyle(workbook);
        styleInitializer.accept(newCellStyle);
        ConditionStyle conditionStyle = new ConditionStyle(condition, newCellStyle);

        /*保存映射结果*/
        if (!colIndex2Styles.containsKey(colIndex)) {
            colIndex2Styles.put(colIndex, new ArrayList<>());
        }
        colIndex2Styles.get(colIndex).add(conditionStyle);
    }


    public static class DefaultStyles {

        DefaultStyles(Workbook workbook) {

        }

        public static final String
    }
}

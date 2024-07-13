package cn.creekmoon.excel.core.R;

import cn.creekmoon.excel.core.R.reader.Reader;
import cn.creekmoon.excel.core.R.reader.ReaderContext;
import cn.creekmoon.excel.core.R.reader.cell.CellReader;
import cn.creekmoon.excel.core.R.reader.cell.HutoolCellReader;
import cn.creekmoon.excel.core.R.reader.title.HutoolTitleReader;
import cn.creekmoon.excel.core.R.reader.title.TitleReader;
import cn.creekmoon.excel.core.R.readerResult.ReaderResult;
import cn.creekmoon.excel.core.R.readerResult.cell.CellReaderResult;
import cn.creekmoon.excel.core.R.readerResult.title.TitleReaderResult;
import cn.creekmoon.excel.util.ExcelConstants;
import cn.creekmoon.excel.util.ExcelFileUtils;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.lang.UUID;
import cn.hutool.core.map.BiMap;
import cn.hutool.core.text.csv.CsvReader;
import cn.hutool.core.text.csv.CsvRow;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.BigExcelWriter;
import cn.hutool.poi.excel.ExcelUtil;
import cn.hutool.poi.excel.sax.Excel07SaxReader;
import cn.hutool.poi.excel.sax.handler.RowHandler;
import jakarta.servlet.http.HttpServletResponse;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.BufferedOutputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStreamReader;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.Supplier;

/**
 * @author JY
 * @date 2022-01-05
 */
@Slf4j
public class ExcelImport {

    public BiMap<Integer, Reader> sheetIndex2ReaderBiMap = new BiMap<>(new HashMap<>());
    public HashMap<Integer, ReaderContext> sheetIndex2ReaderContext = new HashMap<>();
    public HashMap<Integer, ReaderResult> sheetIndex2ReadResult = new HashMap<>();

    /*唯一识别名称 会同步生成一份文件到临时目录*/
    public String taskId = UUID.fastUUID().toString();

    /*当前导入的文件*/
    public MultipartFile sourceFile;


    private ExcelImport() {
    }


    public static ExcelImport create(MultipartFile file) throws IOException {
        ExcelImport excelImport = new ExcelImport();
        excelImport.sourceFile = file;
        excelImport.csvSupport();

        return excelImport;
    }


    /**
     * 获取SHEET页的总行数
     *
     * @return
     */
    @SneakyThrows
    public Long getSheetRowCount(int targetSheetIndex) {
        AtomicLong result = new AtomicLong(0);
        Excel07SaxReader excel07SaxReader = new Excel07SaxReader(new RowHandler() {
            @Override
            public void handle(int sheetIndex, long rowIndex, List<Object> rowCells) {
                if (sheetIndex != targetSheetIndex) {
                    return;
                }
                result.incrementAndGet();
            }
        });
        try {
            excel07SaxReader.read(this.sourceFile.getInputStream(), -1);
        } catch (Exception e) {
            log.error("getSheetRowCount方法读取文件异常", e);
        }
        return result.get();
    }


    /**
     * 切换读取的sheet页
     *
     * @param sheetIndex 下标,从0开始
     * @param supplier   按行读取时,每行数据的实例化对象构造函数
     * @param <T>
     * @return
     */
    public <T> TitleReader<T> switchSheet(int sheetIndex, Supplier<T> supplier) {

        //如果已经存在读取器
        TitleReader sheetReader = (TitleReader) this.sheetIndex2ReaderBiMap.get(sheetIndex);
        if (sheetReader != null) {
            return sheetReader;
        }

        //新增读取器
        HutoolTitleReader<T> reader = new HutoolTitleReader<>(this);
        this.sheetIndex2ReaderBiMap.put(sheetIndex, reader);
        this.sheetIndex2ReaderContext.put(sheetIndex, new ReaderContext(sheetIndex, supplier));
        this.sheetIndex2ReadResult.put(sheetIndex, new TitleReaderResult<T>());
        return reader;
    }


    /**
     * 切换读取的sheet页
     *
     * @param sheetIndex 下标,从0开始
     * @param supplier   按行读取时,每行数据的实例化对象构造函数
     * @param <T>
     * @return
     */
    public <T> CellReader<T> switchSheetAndUseCellReader(int sheetIndex, Supplier<T> supplier) {
        //如果已经存在读取器
        CellReader sheetReader = (CellReader) this.sheetIndex2ReaderBiMap.get(sheetIndex);
        if (sheetReader != null) {
            return sheetReader;
        }

        //新增读取器
        HutoolCellReader<T> reader = new HutoolCellReader<>(this);
        this.sheetIndex2ReaderBiMap.put(sheetIndex, reader);
        this.sheetIndex2ReaderContext.put(sheetIndex, new ReaderContext(sheetIndex, supplier));
        this.sheetIndex2ReadResult.put(sheetIndex, new CellReaderResult<T>());
        return reader;
    }


    /**
     * 支持csv类型的文件  本质是内部将csv转为xlsx
     *
     * @return
     */
    @SneakyThrows
    protected ExcelImport csvSupport() {

        if (StrUtil.isBlank(this.sourceFile.getOriginalFilename()) || !this.sourceFile.getOriginalFilename().toLowerCase().contains(".csv")) {
            /*如果不是csv文件 跳过这个方法*/
            return this;
        }
        log.info("[文件导入]收到CSV格式的文件[{}],尝试转化为XLSX格式", this.sourceFile.getOriginalFilename());
        /*获取CSV文件并尝试读取*/
        String csvTaskId = UUID.fastUUID().toString(true);
        BigExcelWriter bigWriter = ExcelUtil.getBigWriter(ExcelFileUtils.getAbsoluteFilePath(csvTaskId));
        try {
            CsvReader read = new CsvReader(new InputStreamReader(sourceFile.getInputStream()), null);
            Iterator<CsvRow> rowIterator = read.iterator();
            while (rowIterator.hasNext()) {
                bigWriter.writeRow(rowIterator.next().getRawList());
            }
        } catch (Exception e) {
            log.error("[文件导入]csv转换异常!", e);
            this.sourceFile = null;
        } finally {
            bigWriter.close();
            ExcelFileUtils.cleanTempFileDelay(csvTaskId, 120);
        }

        /*将新的xlsx文件替换为当前的文件*/
        this.sourceFile = new MockMultipartFile("csv2xlsx.xlsx", FileUtil.getInputStream(ExcelFileUtils.getAbsoluteFilePath(csvTaskId)));
        return this;
    }


    public ExcelImport response(HttpServletResponse response) throws IOException {
        File file = generateResultFile();
//        IoUtil.copy(sourceFile.getInputStream(), FileUtil.getOutputStream(ExcelFileUtils.getAbsoluteFilePath(this.taskId)));
//        File file = FileUtil.file(ExcelFileUtils.getAbsoluteFilePath(this.taskId));
//        System.out.println("file.canRead() = " + file.canRead());
//        System.out.println("file.canWrite() = " + file.canWrite());
        ExcelFileUtils.response(ExcelFileUtils.getAbsoluteFilePath(this.taskId), taskId + ".xlsx", response);
        return this;
    }

    /**
     * 生成导入结果
     *
     * @return taskId
     */
    public File generateResultFile() throws IOException {
        return this.generateResultFile(true);
    }


    /**
     * 生成导入结果
     *
     * @param autoClean 是否自动删除临时文件(后台进行延迟删除)
     * @return File  生成的新结果文件
     */
    public File generateResultFile(boolean autoClean) throws IOException {

        String absoluteFilePath = ExcelFileUtils.getAbsoluteFilePath(taskId);

        try (Workbook workbook = new XSSFWorkbook(sourceFile.getInputStream());
             BufferedOutputStream outputStream = FileUtil.getOutputStream(absoluteFilePath)) {
            for (Integer targetSheetIndex : sheetIndex2ReaderBiMap.keySet()) {
                Sheet sheet = workbook.getSheetAt(targetSheetIndex);
                Reader reader = sheetIndex2ReaderBiMap.get(targetSheetIndex);
                if (reader instanceof TitleReader titleReader) {

                    //拿上下文状态
                    TitleReaderResult readerResult = (TitleReaderResult) sheetIndex2ReadResult.get(targetSheetIndex);
                    ReaderContext readerContext = titleReader.getReaderContext();

                    // 推算准备要写的位置
                    int titleRowIndex = readerContext.titleRowIndex;
                    Integer lastTitleColumnIndex = readerContext.getLastTitleColumnIndex();
                    int msgTitleColumnIndex = lastTitleColumnIndex + 1;
                    Integer dataFirstRowIndex = readerResult.getDataFirstRowIndex();
                    Integer dataLatestRowIndex = readerResult.getDataLatestRowIndex();

                    // 开始写结果行
                    CellStyle titleCellStyle = sheet.getRow(titleRowIndex).getCell(lastTitleColumnIndex).getCellStyle();
                    Cell cell1 = sheet.getRow(titleRowIndex).createCell(msgTitleColumnIndex);
                    cell1.setCellStyle(titleCellStyle);
                    cell1.setCellValue(ExcelConstants.RESULT_TITLE);

                    // 设置导入结果内容
                    for (Integer rowIndex = dataFirstRowIndex; rowIndex <= dataLatestRowIndex; rowIndex++) {
                        Cell cell = sheet.getRow(rowIndex).createCell(msgTitleColumnIndex);
                        cell.setCellValue(readerResult.getResultMsg(rowIndex));
                    }
                }
            }
            workbook.write(outputStream);
            outputStream.flush();
        }
        return FileUtil.file(absoluteFilePath);
    }

}

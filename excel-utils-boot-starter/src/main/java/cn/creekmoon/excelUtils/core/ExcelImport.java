package cn.creekmoon.excelUtils.core;

import cn.creekmoon.excelUtils.config.ExcelUtilsConfig;
import cn.creekmoon.excelUtils.core.reader.ICellReader;
import cn.creekmoon.excelUtils.core.reader.IReader;
import cn.creekmoon.excelUtils.core.reader.ITitleReader;
import cn.creekmoon.excelUtils.threadPool.CleanTempFilesExecutor;
import cn.hutool.core.io.FileUtil;
import cn.hutool.core.text.csv.CsvReader;
import cn.hutool.core.text.csv.CsvRow;
import cn.hutool.core.util.StrUtil;
import cn.hutool.poi.excel.sax.Excel07SaxReader;
import cn.hutool.poi.excel.sax.handler.RowHandler;
import jakarta.servlet.http.HttpServletResponse;
import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStreamReader;
import java.util.*;
import java.util.concurrent.Semaphore;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicLong;
import java.util.function.Supplier;

/**
 * @author JY
 * @date 2022-01-05
 */
@Slf4j
public class ExcelImport {
    /**
     * 控制导入并发数量
     */
    public static Semaphore importSemaphore;

    public HashMap<Integer, IReader> sheetIndex2Reader = new HashMap<>();
    public HashMap<Integer, ReaderResult> sheetIndex2ReadResult = new HashMap<>();

    @Deprecated
    public HashMap<Object, HashMap<String, Object>> convertObject2rawData = new HashMap<>();
    @Deprecated
    public HashMap<Integer, List<Map<String, Object>>> sheetIndex2rawData = new LinkedHashMap<>();


    /*当前导入的文件*/
    protected MultipartFile file;

    /*错误次数统计*/
    @Deprecated
    protected AtomicInteger errorCount = new AtomicInteger(0);


    private ExcelImport() {
    }


    public static void init() {
        if (importSemaphore == null) {
            importSemaphore = new Semaphore(ExcelUtilsConfig.IMPORT_MAX_PARALLEL);
        }
    }


    public static ExcelImport create(MultipartFile file) {
        if (importSemaphore == null) {
            throw new RuntimeException("请使用@EnableExcelUtils进行初始化配置!");
        }
        ExcelImport excelImport = new ExcelImport();
        excelImport.file = file;
        excelImport.csvSupport();


        return excelImport;
    }


    public static <T> IReader<T> create(MultipartFile file, Supplier<T> supplier) {
        ExcelImport excelImport = create(file);
        return excelImport.switchSheet(0, supplier);
    }

    /**
     * 切换读取的sheet页
     *
     * @param sheetIndex 下标,从0开始
     * @param supplier   按行读取时,每行数据的实例化对象构造函数
     * @param <T>
     * @return
     */
    public <T> ICellReader<T> switchSheetAndUseCellReader(int sheetIndex, Supplier<T> supplier) {
        return HutoolCellReader.of(new ReaderContext(sheetIndex, supplier), this);
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
            excel07SaxReader.read(this.file.getInputStream(), -1);
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
    public <T> ITitleReader<T> switchSheet(int sheetIndex, Supplier<T> supplier) {
        ITitleReader sheetReader = (ITitleReader) this.sheetIndex2Reader.get(sheetIndex);
        if (sheetReader != null) {
            return sheetReader;
        }

        ReaderContext context = new ReaderContext(sheetIndex, supplier);

        HutoolTitleReader<T> reader = new HutoolTitleReader<>();
        reader.readerContext = context;
        reader.parent = this;


        sheetIndex2Reader.put(reader.readerContext.sheetIndex, reader);
        return reader;
    }


    /**
     * 支持csv类型的文件  本质是内部将csv转为xlsx
     *
     * @return
     */
    @SneakyThrows
    protected ExcelImport csvSupport() {

        if (StrUtil.isBlank(this.file.getOriginalFilename()) || !this.file.getOriginalFilename().toLowerCase().contains(".csv")) {
            /*如果不是csv文件 跳过这个方法*/
            return this;
        }

        /*获取CSV文件并尝试读取*/
        ExcelExport newExcel = ExcelExport.create();
        try {
            CsvReader read = new CsvReader(new InputStreamReader(file.getInputStream()), null);
            Iterator<CsvRow> rowIterator = read.iterator();
            while (rowIterator.hasNext()) {
                newExcel.getBigExcelWriter().writeRow(rowIterator.next().getRawList());
            }
        } catch (Exception e) {
            log.error("csv转换异常!", e);
            e.printStackTrace();
            this.file = null;
        } finally {
            CleanTempFilesExecutor.cleanTempFileDelay(newExcel.stopWrite(), 10);
        }

        /*将新的xlsx文件替换为当前的文件*/
        this.file = new MockMultipartFile("csv2xlsx.xlsx", FileUtil.getInputStream(PathFinder.getAbsoluteFilePath(newExcel.taskId)));
        return this;
    }


    public ExcelImport response(HttpServletResponse response) throws IOException {
        Workbook workbook = new XSSFWorkbook(file.getInputStream());

        for (Integer targetSheetIndex : sheetIndex2Reader.keySet()) {
            IReader reader = sheetIndex2Reader.get(targetSheetIndex);
            if (reader instanceof ITitleReader) {
                ITitleReader titleReader = (ITitleReader) reader;
                TitleReaderResult readerResult = (TitleReaderResult) sheetIndex2ReadResult.get(targetSheetIndex);
                ReaderContext readerContext = reader.getReaderContext();
                Integer resultColIndex = readerContext.colIndex2Title.keySet().stream().max(Integer::compareTo).get() + 1;


                Sheet sheetAt = workbook.getSheetAt(targetSheetIndex);
                Row row = sheetAt.getRow(readerContext.firstRowIndex);
                Cell cell = row.createCell(resultColIndex);
                cell.setCellValue();


            }

        }


        ExcelExport.response(generateResultFile(), excelExport.excelName, response);
        return this;
    }

    /**
     * 生成导入结果
     *
     * @return taskId
     */
    public String generateResultFile() {
        return this.generateResultFile(true);
    }


    /**
     * 生成导入结果
     *
     * @param autoCleanTempFile 自动删除临时文件(后台进行延迟删除)
     * @return taskId
     */
    public String generateResultFile(boolean autoCleanTempFile) {
        if (!sheetIndex2rawData.isEmpty()) {
            sheetIndex2rawData.forEach((k, v) -> {
                excelExport
                        .switchSheet(ExcelExport.generateSheetNameByIndex(k), Map.class)
                        .setColumnWidthDefault()
                        .writeByMap(v);
            });
        }
        String taskId = excelExport.stopWrite();
        if (autoCleanTempFile) {
            CleanTempFilesExecutor.cleanTempFileDelay(taskId, 10);
        }
        return taskId;
    }

    /**
     * 设置读取的结果
     *
     * @param object 实例化的对象
     * @param msg    结果
     */
    public void setResult(Object object, String msg) {
        Map<String, Object> row = convertObject2rawData.get(object);
        if (row != null) {
            row.put(ExcelConstants.RESULT_TITLE, msg);
        }
    }

    public static void cleanTempFileDelay(String taskId) {
        CleanTempFilesExecutor.cleanTempFileDelay(taskId);
    }

    @Deprecated
    public AtomicInteger getErrorCount() {
        return errorCount;
    }
}

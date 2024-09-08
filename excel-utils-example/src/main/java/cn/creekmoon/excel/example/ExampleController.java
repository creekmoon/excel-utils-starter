package cn.creekmoon.excel.example;

import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.DateConverter;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.R.converter.LocalDateTimeConverter;
import cn.creekmoon.excel.core.R.readerResult.ReaderResult;
import cn.creekmoon.excel.core.R.readerResult.title.TitleReaderResult;
import cn.creekmoon.excel.core.W.ExcelExport;
import cn.hutool.core.util.RandomUtil;
import io.swagger.v3.oas.annotations.Operation;
import io.swagger.v3.oas.annotations.tags.Tag;
import jakarta.servlet.http.HttpServletRequest;
import jakarta.servlet.http.HttpServletResponse;
import lombok.extern.slf4j.Slf4j;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;

@Tag(name = "测试API")
@RestController
@Slf4j
public class ExampleController {


    @GetMapping(value = "/exportExcel")
    @Operation(summary = "导出")
    public void exportExcel(Integer size, HttpServletRequest request, HttpServletResponse response) throws IOException {
        /*查询数据*/
        ArrayList<Student> result = createStudentList(size != null ? size : 60_000);
        ArrayList<Student> result2 = createStudentList(size != null ? size : 10_000);

        ExcelExport excelExport = ExcelExport.create();
        excelExport
                .switchNewSheet(Student.class)
                /*第一个标签页*/
                .addTitle("基本信息::用户名", Student::getUserName)
                .addTitle("基本信息::全名", Student::getFullName)
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .addTitle("生日", Student::getBirthday)
                .addTitle("过期时间", Student::getExpTime)
                .write(result);

        /*第二个标签页*/
        excelExport.switchNewSheet(Student.class)
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .addTitle("生日", Student::getBirthday)
                .addTitle("过期时间", Student::getExpTime)
                .write(result2)
                .response(response);
    }

    @GetMapping(value = "/exportExcelByStyle")
    @Operation(summary = "导出(并设置style)")
    public void exportExcel5(HttpServletRequest request, HttpServletResponse response) throws IOException {
//        ArrayList<Student> result = createStudentList(60_000);
//        ArrayList<Teacher> result2 = createTeacherList(60_000);
//
//        ExcelExport excelExport = ExcelExport.create();
//
//        /*定义一个全局的数据样式  double是千分号和保留两位小数  int是千分号,保留整数*/
//        short dataFormat_double = excelExport.getBigExcelWriter().getWorkbook().createDataFormat().getFormat("#,##0.00");
//        short dataFormat_int = excelExport.getBigExcelWriter().getWorkbook().createDataFormat().getFormat("#,##0");
//
//        TitleWriter<Student> studentTitleWriter = excelExport.switchNewSheet(Student.class);
//        studentTitleWriter
//                .addTitle("用户名", Student::getUserName)
//                .addTitle("全名", Student::getFullName)
//                .setDataStyle(cellStyle ->
//                {
//                    cellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
//                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//                })
//                .addTitle("年龄", Student::getAge)
//                .setDataStyle(student -> student.getAge() > 20,
//                        cellStyle ->
//                        {
//                            cellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
//                            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
//                        }
//                )
//                .setDataStyle(student -> student.getAge() > 20,
//                        cellStyle ->
//                        {
//                            cellStyle.setDataFormat(dataFormat_double);
//                        }
//                )
//                .write(result)
//                .response(response);
    }

    /**
     * 导入
     *
     * @param request
     * @param response
     * @throws IOException
     */
    @PostMapping(value = "/importExcel")
    @Operation(summary = "导入数据")
    public void importExcelBySax(MultipartFile file, HttpServletRequest request, HttpServletResponse response) throws Exception {
        /*读取导入文件*/
        ExcelImport excelImport = ExcelImport.create(file);
        TitleReaderResult<Student> sheet1 = excelImport.switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .read(x -> log.info(x.toString()));

        TitleReaderResult<Student> sheet2 = excelImport.switchSheet(1, Student::new)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("邮箱", Student::setEmail)
                .read(x -> {
                    if ("lisi@qq.com".equals(x.getEmail())) {
                        Thread.sleep(2000);
                        throw new RuntimeException("错误!");
                    }
                    log.info(x.toString());
                });

        excelImport.response(response);
    }


    /**
     * 导入
     *
     * @param request
     * @param response
     * @throws IOException
     */
    @PostMapping(value = "/importExcelByCell")
    @Operation(summary = "导入数据(读取指定单元格)")
    public void importExcelByCell(MultipartFile file, HttpServletRequest request, HttpServletResponse response) throws Exception {

        ExcelImport excelImport = ExcelImport.create(file);
        ReaderResult<Student> readerResult = excelImport.switchSheetAndUseCellReader(1, Student::new)
                .addConvertAndMustExist("A1", Student::setUserName)
                .read(x -> {
                    log.info(x.toString());
                });

        System.out.println("readerResult.getErrorReport() = " + readerResult.errorReport);
        excelImport.response(response);

    }


    private Student createNewStudent() {
        Student student = new Student();
        //随机年龄
        student.setAge(RandomUtil.randomInt(1, 100));
        student.setBirthday(new Date());
        //随机生成邮箱
        student.setEmail(RandomUtil.randomString(10) + "@qq.com");
        //随机生成时间
        student.setExpTime(LocalDateTime.now());
        student.setFullName(RandomUtil.randomString(5));
        student.setUserName(RandomUtil.randomString(5));
        student.setBirthday(new Date());
        return student;
    }

    private Teacher createNewTeacher() {
        Teacher teacher = new Teacher();
        //随机年龄
        teacher.setWorkYear(RandomUtil.randomInt(1, 10));
        teacher.setTeacherName(RandomUtil.randomString(5));
        return teacher;
    }

    private ArrayList<Student> createStudentList(int size) {
        ArrayList<Student> result = new ArrayList<>();
        //加入数据
        for (int i = 0; i < size; i++) {
            result.add(createNewStudent());
        }
        return result;
    }

    private ArrayList<Teacher> createTeacherList(int size) {
        ArrayList<Teacher> result = new ArrayList<>();
        //加入数据
        for (int i = 0; i < size; i++) {
            result.add(createNewTeacher());
        }
        return result;
    }


}

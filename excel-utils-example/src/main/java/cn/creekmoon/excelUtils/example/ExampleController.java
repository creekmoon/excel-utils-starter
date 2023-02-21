package cn.creekmoon.excelUtils.example;

import cn.creekmoon.excelUtils.converter.DateConverter;
import cn.creekmoon.excelUtils.converter.IntegerConverter;
import cn.creekmoon.excelUtils.converter.LocalDateTimeConverter;
import cn.creekmoon.excelUtils.core.AsyncTaskState;
import cn.creekmoon.excelUtils.core.ExcelConstants;
import cn.creekmoon.excelUtils.core.ExcelExport;
import cn.creekmoon.excelUtils.core.ExcelImport;
import cn.creekmoon.excelUtils.example.config.exception.MyNewException;
import cn.creekmoon.excelUtils.hutool589.core.text.StrFormatter;
import cn.hutool.core.util.RandomUtil;
import io.swagger.annotations.Api;
import io.swagger.annotations.ApiOperation;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.PostMapping;
import org.springframework.web.bind.annotation.RestController;
import org.springframework.web.multipart.MultipartFile;

import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

@Api(tags = "测试API")
@RestController
@Slf4j
public class ExampleController {

    // key=taskId  value=异步状态  这里模拟保存到redis中
    private static final Map<String, AsyncTaskState> taskId2TaskState = new ConcurrentHashMap<>();

    @GetMapping(value = "/exportExcel")
    @ApiOperation("单次查询,并导出数据")
    public void exportExcel(Integer size, HttpServletRequest request, HttpServletResponse response) throws IOException {
        ArrayList<Student> result = createStudentList(size != null ? size : 60_000);
        ExcelExport.create(StrFormatter.format("导出数据"), Student.class)
                .addTitle("用户名", Student::getUserName)
                .addTitle("全名", Student::getFullName)
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .addTitle("生日", Student::getBirthday)
                .addTitle("过期时间", Student::getExpTime)
                .write(result)
                .response(response);
    }


    @GetMapping(value = "/exportExcel2")
    @ApiOperation("多次查询,并导出数据")
    public void exportExcel2(HttpServletRequest request, HttpServletResponse response) throws IOException {

        /*构建表头*/
        ExcelExport<Student> excelExport = ExcelExport.create("lalala", Student.class)
                .addTitle("用户名", Student::getUserName)
                .addTitle("全名", Student::getFullName)
                .addTitle("年龄", Student::getAge)
                .addTitle("邮箱", Student::getEmail)
                .addTitle("生日", Student::getBirthday)
                .addTitle("过期时间", Student::getExpTime);
        //模拟查询
        for (int i = 0; i < 3; i++) {
            excelExport.write(createStudentList(250_000));
        }
        /*返回数据*/
        excelExport.response(response);
    }


    @GetMapping(value = "/exportExcel3")
    @ApiOperation("构建多级表头,导出数据")
    public void exportExcel3(HttpServletRequest request, HttpServletResponse response) throws IOException {
        ArrayList<Student> result = createStudentList(60_000);
        ExcelExport.create("lalala", Student.class)
                .addTitle("基本信息::用户名", Student::getUserName)
                .addTitle("基本信息::全名", Student::getFullName)
                .addTitle("额外附加信息::年龄", Student::getAge)
                .addTitle("额外附加信息::邮箱", Student::getEmail)
                .addTitle("额外附加信息::系统数据::生日", Student::getBirthday)
                .addTitle("额外附加信息::系统数据::过期时间", Student::getExpTime)
                .write(result)
                .response(response);
    }

    @GetMapping(value = "/exportExcel4")
    @ApiOperation("构建多个Sheet页,导出数据")
    public void exportExcel4(HttpServletRequest request, HttpServletResponse response) throws IOException {
        ArrayList<Student> result = createStudentList(60_000);
        ExcelExport.create("lalala", Student.class)
                .switchSheet("第一个标签页", Student.class)
                .addTitle("基本信息::用户名", Student::getUserName)
                .addTitle("基本信息::全名", Student::getFullName)
                .write(result)
                .switchSheet("第二个标签页", Student.class)
                .addTitle("额外附加信息::年龄", Student::getAge)
                .addTitle("额外附加信息::邮箱", Student::getEmail)
                .addTitle("额外附加信息::系统数据::生日", Student::getBirthday)
                .addTitle("额外附加信息::系统数据::过期时间", Student::getExpTime)
                .write(result)
                .response(response);
    }


    @GetMapping(value = "/exportExcel5")
    @ApiOperation("构建多个Sheet页,导出数据,并设置style")
    public void exportExcel5(HttpServletRequest request, HttpServletResponse response) throws IOException {
        ArrayList<Student> result = createStudentList(60_000);
        ExcelExport.create("lalala", Student.class)
                .switchSheet("第一个标签页", Student.class)
                .addTitle("基本信息::用户名", Student::getUserName)
                .addTitle("基本信息::全名(全部标黄)", Student::getFullName)
                .setDataStyle(cellStyle ->
                {
                    cellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
                    cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                })
                .write(result)
                .switchSheet("第二个标签页", Student.class)
                .addTitle("额外附加信息::年龄(大于25标黄)", Student::getAge)
                .setConditionDataStyle(student -> student.getAge() > 25,
                        cellStyle ->
                        {
                            cellStyle.setFillForegroundColor(IndexedColors.LIGHT_ORANGE.getIndex());
                            cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                        }
                )
                .addTitle("额外附加信息::邮箱", Student::getEmail)
                .addTitle("额外附加信息::系统数据::生日", Student::getBirthday)
                .addTitle("额外附加信息::系统数据::过期时间", Student::getExpTime)
                .write(result)
                .response(response);
    }

    /**
     * 导入
     *
     * @param request
     * @param response
     * @throws IOException
     */
    @PostMapping(value = "/importExcelBySax")
    @ApiOperation("同步导入数据(SAX模式)")
    public void importExcelBySax(MultipartFile file, HttpServletRequest request, HttpServletResponse response) throws IOException {
        //判断这个方法的执行时间
        long start = System.currentTimeMillis();
        ExcelImport.create(file, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .read(student -> {
                    if (student.age == 76) {
                        throw new MyNewException("年龄==76");
                    }
                    System.out.println(student);
                })
                .response(response);

        //判断这个方法的执行时间
        long end = System.currentTimeMillis();
        System.out.println("执行时间:" + (end - start));
    }

    /**
     * 导入
     *
     * @param request
     * @param response
     * @throws IOException
     */
    @PostMapping(value = "/importExcelByMemory")
    @ApiOperation("同步导入数据(内存模式)")
    public void importExcelByMemory(MultipartFile file, HttpServletRequest request, HttpServletResponse response) throws IOException {

        ExcelImport<Student> studentExcelImport = ExcelImport.create(file, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime);

        //把excel都读取到内存中
        List<Student> students = studentExcelImport.readAll(ExcelImport.ConvertStrategy.SKIP_ALL_IF_FAIL);
        students.forEach(student -> {
            if (student.age > 50) {
                studentExcelImport.setResult(student, "学生年龄大于50, 不合法, 请检查参数");
                return;
            }
            //todo 执行一些业务操作
            studentExcelImport.setResult(student, ExcelConstants.IMPORT_SUCCESS_MSG);
        });

        //响应导出结果
        studentExcelImport.response(response);
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

    private ArrayList<Student> createStudentList(int size) {
        ArrayList<Student> result = new ArrayList<>();
        //加入数据 六十万
        for (int i = 0; i < size; i++) {
            result.add(createNewStudent());
        }
        return result;
    }


}

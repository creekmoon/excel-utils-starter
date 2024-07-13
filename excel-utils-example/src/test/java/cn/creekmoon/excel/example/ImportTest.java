package cn.creekmoon.excel.example;


import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.DateConverter;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.R.converter.LocalDateTimeConverter;
import cn.creekmoon.excel.core.R.readerResult.title.TitleReaderResult;
import cn.hutool.core.io.resource.ResourceUtil;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.autoconfigure.web.servlet.AutoConfigureMockMvc;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.mock.web.MockHttpSession;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.test.web.servlet.MockMvc;

import java.io.IOException;
import java.io.InputStream;

@AutoConfigureMockMvc
@SpringBootTest
@Slf4j
public class ImportTest {

    @Autowired
    private MockMvc mvc;
    private MockHttpSession session;

    @BeforeEach
    public void init() {
        session = new MockHttpSession();
    }

    @Test
    void importTest() throws IOException, InterruptedException {
        /*读取导入文件*/
        String IMPORT_FILE_NAME = "import-demo-1000.xlsx";
        InputStream stream = ResourceUtil.getStream(IMPORT_FILE_NAME);
        MockMultipartFile mockMultipartFile = new MockMultipartFile(IMPORT_FILE_NAME, stream);

        ExcelImport excelImport = ExcelImport.create(mockMultipartFile);
        TitleReaderResult<Student> read = excelImport.switchSheet(0, Student::new)
                .addConvert("用户名", Student::setUserName)
                .addConvert("全名", Student::setFullName)
                .addConvert("年龄", IntegerConverter::parse, Student::setAge)
                .addConvert("邮箱", Student::setEmail)
                .addConvert("生日", DateConverter::parse, Student::setBirthday)
                .addConvert("过期时间", LocalDateTimeConverter::parse, Student::setExpTime)
                .read();
        read.consume(x -> {
            if (x.age == null || x.age > 60) {
                throw new RuntimeException("学生年龄过大!");
            }
        });
        System.out.println("read.getDurationSecond() = " + read.getDurationSecond());
        System.out.println("read.getErrorCount() = " + read.getErrorCount());
    }

}

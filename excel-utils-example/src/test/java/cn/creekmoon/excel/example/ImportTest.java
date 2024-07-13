package cn.creekmoon.excel.example;


import cn.creekmoon.excel.core.R.ExcelImport;
import cn.creekmoon.excel.core.R.converter.DateConverter;
import cn.creekmoon.excel.core.R.converter.IntegerConverter;
import cn.creekmoon.excel.core.R.converter.LocalDateTimeConverter;
import cn.creekmoon.excel.core.R.readerResult.title.TitleReaderResult;
import cn.hutool.core.io.resource.ResourceUtil;
import lombok.extern.slf4j.Slf4j;
import org.junit.jupiter.api.Assertions;
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
import java.util.List;

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
        List<Student> dataList = read.getAll();
        // 数据从第二行开始, 而索引下标从0开始, 所以需要都-2
        Assertions.assertEquals(88, dataList.get(6 - 2).age, "第6行数据年龄为88");
        Assertions.assertEquals("poxo0", dataList.get(80 - 2).getUserName(), "第80行数据用户名是poxo0");
        Assertions.assertEquals("npdsytu8i5@qq.com", dataList.get(195 - 2).getEmail(), "第195行数据邮箱是npdsytu8i5@qq.com");
        //第630条全名是j94rk
        Assertions.assertEquals("j94rk", dataList.get(631 - 2).getFullName(), "第631行数据全名是j94rk");
        //全部数据有1000条
        Assertions.assertEquals(1000, dataList.size(), "全部数据有1000条");
        Assertions.assertEquals(1, read.getDataFirstRowIndex(), "数据起始行下标预期为1");
        Assertions.assertEquals(1000, read.getDataLatestRowIndex(), "数据结束行下标预期为1000");

        long countAgeLg60 = dataList
                .stream()
                .map(Student::getAge)
                .filter(x -> x > 60)
                .count();
        Assertions.assertEquals(countAgeLg60, 411, "年龄大于60应该为411个");


    }

}

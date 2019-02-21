package com.proaim;

import com.proaim.entity.Model;
import com.sargeraswang.util.ExcelUtil.ExcelLogs;
import com.sargeraswang.util.ExcelUtil.ExcelUtil;
import org.junit.Test;
import org.junit.runner.RunWith;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.test.context.junit4.SpringRunner;

import java.io.*;
import java.util.*;

@RunWith(SpringRunner.class)
@SpringBootTest
public class ExcelUtilApplicationTests {

    @Test
    // 导出-Bean方式
    public void exportBeanXls() throws IOException {
        //用排序的Map且Map的键应与ExcelCell注解的index对应
        Map<String, String> map = new LinkedHashMap<>();
        map.put("a", "姓名");
        map.put("b", "年龄");
        map.put("c", "性别");
        map.put("d", "出生日期");
        Collection<Object> dataset = new ArrayList<Object>();
        dataset.add(new Model("", "", "", null));
        dataset.add(new Model(null, null, null, null));
        dataset.add(new Model("王五", "34", "男", new Date()));
        File file = new File(new File("").getCanonicalPath() + "/src/main/resources/excel/test.xls");
        OutputStream out = new FileOutputStream(file);
        ExcelUtil.exportExcel(map, dataset, out);
        out.close();
    }

    @Test
    // 导出-Map方式
    public void exportMapXls() throws IOException {
        List<Map<String, Object>> list = new ArrayList<>();
        Map<String, Object> map = new LinkedHashMap<>();
        map.put("name", "");
        map.put("age", "");
        map.put("birthday", "");
        map.put("sex", "");
        Map<String, Object> map2 = new LinkedHashMap<String, Object>();
        map2.put("name", "测试是否是中文长度不能自动宽度.测试是否是中文长度不能自动宽度.");
        map2.put("age", null);
        map2.put("sex", null);
        map.put("birthday", null);
        Map<String, Object> map3 = new LinkedHashMap<String, Object>();
        map3.put("name", "张三");
        map3.put("age", 12);
        map3.put("sex", "男");
        map3.put("birthday", new Date());
        list.add(map);
        list.add(map2);
        list.add(map3);
        Map<String, String> map1 = new LinkedHashMap<>();
        map1.put("name", "姓名");
        map1.put("age", "年龄");
        map1.put("birthday", "出生日期");
        map1.put("sex", "性别");
        File file = new File(new File("").getCanonicalPath() + "/src/main/resources/excel/test.xls");
        OutputStream out = new FileOutputStream(file);
        ExcelUtil.exportExcel(map1, list, out);
        out.close();
    }


    // xls导入
    @Test
    public void importXls() throws FileNotFoundException {
        File f = new File("src/main/resources/excel/test.xls");
        InputStream inputStream = new FileInputStream(f);

        ExcelLogs logs = new ExcelLogs();
        Collection<Map> importExcel = ExcelUtil.importExcel(Map.class, inputStream, "yyyy/MM/dd HH:mm:ss", logs, 0);

        for (Map m : importExcel) {
            System.out.println(m);
        }
    }

    // xlsx导入
    @Test
    public void importXlsx() throws FileNotFoundException {
        File file = new File("src/main/resources/excel/test.xlsx");
        InputStream inputStream = new FileInputStream(file);

        ExcelLogs logs = new ExcelLogs();
        Collection<Map> importExcel = ExcelUtil.importExcel(Map.class, inputStream, "yyyy/MM/dd HH:mm:ss", logs, 0);

        for (Map m : importExcel) {
            System.out.println(m);
        }
    }

}

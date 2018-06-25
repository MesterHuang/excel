package excelTest;

import excelUtil.ExcelUtil;
import excelUtil.Student;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

import static excelUtil.ExcelUtil1.exportUserExcel;

/**
 * @Author: HuangC
 * @Description:
 * @Date: 2018/6/13 17:19
 */
public class Test {


    public static void main(String[] args) {
        List<Student> list = new ArrayList<Student>();
        Student student = new Student();
        student.setName("A");
        student.setAge(23);
        student.setSex("男");
        student.setData(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date()));
        Student student1 = new Student();
        student1.setName("B");
        student1.setAge(23);
        student1.setSex("女");
        student1.setData(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date()));
        Student student2 = new Student();
        student2.setName("c");
        student2.setAge(23);
        student2.setSex("男");
        student2.setData(new SimpleDateFormat("yyyy-MM-dd HH:mm:ss").format(new Date()));
        list.add(student);
        list.add(student1);
        list.add(student2);

        String[] headers = {"姓名","年齡","性別","入學日期"};
        String title = "学生表";
        OutputStream os = null;
        try {
            os = new FileOutputStream("E:"+ File.separator +"234.xls");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        ExcelUtil.exportDataToExcel(list, headers, title, os);
        //exportUserExcel(list,os);
    }
}

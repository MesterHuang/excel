package excelTest;

import excelUtil.User;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import static excelUtil.ExcelUtil1.exportUserExcel;
import static excelUtil.ExcelUtil1.importExcel;

/**
 * @Author: HuangC
 * @Description:
 * @Date: 2018/6/14 14:11
 */
public class Test1 {
    public static void main(String[] args) {
        List<User> list = new ArrayList<User>();
        User user = new User();
        user.setName("A");
        user.setAccount("123");
        user.setDept("部門A");
        user.setGender(true);
        user.setEmail("123@a.com");

        User user1 = new User();
        user1.setName("B");
        user1.setAccount("123");
        user1.setDept("部門B");
        user1.setGender(false);
        user1.setEmail("123@a.com");
        list.add(user);
        list.add(user1);

        OutputStream os = null;
        try {
            os = new FileOutputStream("E:"+ File.separator +"254.xls");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        exportUserExcel(list,os);
        importExcel("E:"+ File.separator+"254.xls");
    }
}

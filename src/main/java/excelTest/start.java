package excelTest;

import java.io.File;

import static excelUtil.ExcelUtil.importDataFromExcel;


/**
 * @Author: HuangC
 * @Description: 导入Excel
 * @Date: 2018/6/13 15:25
 */
public class start {

    public static void main(String[] args)
    {
        importDataFromExcel("E:"+ File.separator +"123.xls");
    }
}

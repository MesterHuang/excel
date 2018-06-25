package excelUtil;

import org.apache.log4j.Logger;
import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;

/**
 * @Author: HuangC
 * @Description:
 * @Date: 2018/6/13 15:20
 */
public class ExcelUtil  {

    private static final Logger logger = Logger.getLogger(ExcelUtil.class);

    /**
     * 將Excel文件导入到數據庫
     * @param filePath 导入文件路径
     */
    public static void importDataFromExcel(String filePath) {

        //判断是否为excel类型文件
        if(!filePath.endsWith(".xls")&&!filePath.endsWith(".xlsx")) {
            System.out.println("文件不是excel类型");
        }

        FileInputStream fis =null;
        Workbook workbook = null;

        try {
            //获取流对象，
            fis = new FileInputStream(filePath);
            //2007版本的excel，用.xlsx结尾

            workbook = new XSSFWorkbook(filePath);//得到工作簿
        } catch (IOException e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        } finally {
            try {
                fis.close();
            } catch (IOException e) {
                e.printStackTrace();
            }
        }

        //得到一个工作表
        Sheet sheet = workbook.getSheetAt(0);

        //获得表头
        Row rowHead = sheet.getRow(0);

        //判断表头是否正确
        if(rowHead.getPhysicalNumberOfCells() != 3) {
            System.out.println("表头的数量不对!");
        }

        //获得数据的总行数
        int totalRowNum = sheet.getLastRowNum();

        /**
         * 方式一
         */
        /*int totalCells = 0;
        // 得到Excel的列数(前提是有行数)
        if (totalRowNum > 1 && sheet.getRow(0) != null) {
            totalCells = sheet.getRow(0).getPhysicalNumberOfCells();
         }

        List<Map<String, Object>> userList = new ArrayList<Map<String, Object>>();
         // 循环Excel行数
         for (int r = 1; r <= totalRowNum; r++) {
             Row row = sheet.getRow(r);
             if (row == null) {
                 continue;
             }
             // 循环Excel的列
             Map<String, Object> map = new HashMap<String, Object>();
             for (int c = 0; c < totalCells; c++) {
                 Cell cell = row.getCell(c);
                 if (null != cell) {
                     if (c == 0) {
                         // 如果是纯数字,比如你写的是25,cell.getNumericCellValue()获得是25.0,通过截取字符串去掉.0获得25
                         if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                             String name = String.valueOf(cell.getNumericCellValue());
                             map.put("name", name.substring(0, name.length() - 2 > 0 ? name.length() - 2 : 1));// 名称
                         } else {
                             map.put("name", cell.getStringCellValue());// 名称
                         }
                     } else if (c == 1) {
                        if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                             String sex = String.valueOf(cell.getNumericCellValue());
                             map.put("sex",sex.substring(0, sex.length() - 2 > 0 ? sex.length() - 2 : 1));// 性别
                         } else {
                             map.put("sex",cell.getStringCellValue());// 性别
                         }
                     } else if (c == 2) {
                         if (cell.getCellType() == HSSFCell.CELL_TYPE_NUMERIC) {
                             String age = String.valueOf(cell.getNumericCellValue());
                             map.put("age", age.substring(0, age.length() - 2 > 0 ? age.length() - 2 : 1));// 年龄
                          } else {
                             map.put("age", cell.getStringCellValue());// 年龄
                         }
                     }
                 }
             }
             // 添加到list
             userList.add(map);

         }
        System.out.println(userList.toString());*/

        /***
         * 方式二
         */
        //要获得属性
        String name = "";
        int age = 0;
        String sex = "";

        Student student = new Student();
        //获得所有数据
        for(int i = 1 ; i <= totalRowNum ; i++) {
            //获得第i行对象
            Row row = sheet.getRow(i);

            //获得获得第i行第0列的 String类型对象
            Cell cell = row.getCell((short)0);
            name = cell.getStringCellValue().toString();
            student.setName(cell.getStringCellValue().toString());
            //获得一个数字类型的数据
            cell = row.getCell((short)1);
            age = (int) cell.getNumericCellValue();
            student.setAge((int)cell.getNumericCellValue());

            cell = row.getCell((short)2);
            sex = cell.getStringCellValue().toString();
            student.setSex(cell.getStringCellValue());

            System.out.println("名字："+name+",年龄："+age+",性别:"+sex);
            System.out.println(student.toString());

        }
    }


    /**
     * 將數據庫的數據導出到Excel
     * @param list 导出的数据
     * @param headers 头
     * @param title 标题
     * @param os 导出路径
     * @param <T>
     */
    public static <T>  void exportDataToExcel(List<T> list, String[] headers, String title, OutputStream os){
        //创建一个工作簿
        HSSFWorkbook workbook = new HSSFWorkbook();
        //生成一个表格
        HSSFSheet sheet = workbook.createSheet(title);
        //设置表格默认列宽15个字节
        sheet.setDefaultColumnWidth(15);
        //生成一个样式
        HSSFCellStyle style = getCellStyle(workbook);
        //生成一个字体
        HSSFFont font = getFont(workbook);
        //把字体应用到当前样式
        style.setFont(font);

        //生成表格标题
        HSSFRow row = sheet.createRow(0);
        row.setHeight((short)300);
        HSSFCell cell = null;

        for (int i = 0; i < headers.length; i++) {
            cell = row.createCell(i);
            cell.setCellStyle(style);
            HSSFRichTextString text = new HSSFRichTextString(headers[i]);
            cell.setCellValue(text);
        }

        //将数据放入sheet中
        for (int i = 0; i < list.size(); i++) {
            row = sheet.createRow(i+1);
            T t = list.get(i);
            //利用反射，根据JavaBean属性的先后顺序，动态调用get方法得到属性的值
            Field[] fields = t.getClass().getDeclaredFields();
            try {
                for (int j = 0; j < fields.length; j++) {
                    cell = row.createCell(j);
                    Field field = fields[j];
                    String fieldName = field.getName();
                    String methodName = "get"+fieldName.substring(0, 1).toUpperCase()+fieldName.substring(1);
                    Method getMethod = t.getClass().getMethod(methodName,new Class[]{});
                    Object value = getMethod.invoke(t, new Object[]{});

                    if(null == value)
                        value ="";
                    cell.setCellValue(value.toString());

                }
            } catch (Exception e) {
                logger.error(e);
            }
        }

        try {
            workbook.write(os);
        } catch (Exception e) {
            logger.error(e);
        }finally{
            try {
                os.flush();
                os.close();
            } catch (IOException e) {
                logger.error(e);
            }
        }

    }

    /**
      * @Title: getCellStyle
      * @Description: 获取单元格格式
      * @param @param workbook
      * @param @return
      * @return HSSFCellStyle
      * @throws
      */
    public static HSSFCellStyle getCellStyle(HSSFWorkbook workbook){
        HSSFCellStyle style = workbook.createCellStyle();
        style.setFillForegroundColor(HSSFColor.BLACK.index);
                 /*style.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
                 style.setBorderBottom(HSSFCellStyle.BORDER_THIN);
                 style.setBorderTop(HSSFCellStyle.BORDER_THIN);
                 style.setLeftBorderColor(HSSFCellStyle.BORDER_THIN);
                 style.setRightBorderColor(HSSFCellStyle.BORDER_THIN);
                 style.setAlignment(HSSFCellStyle.ALIGN_CENTER);*/

        return style;
    }

    /**
     255     * @Title: getFont
     256     * @Description: 生成字体样式
     257     * @param @param workbook
     258     * @param @return
     259     * @return HSSFFont
     260     * @throws
     261     */
    public static HSSFFont getFont(HSSFWorkbook workbook){
        HSSFFont font = workbook.createFont();
        font.setColor(HSSFColor.BLACK.index);
        font.setFontHeightInPoints((short)12);
        //font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        return font;
    }
}

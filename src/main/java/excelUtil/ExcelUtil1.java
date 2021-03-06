package excelUtil;

import org.apache.poi.hssf.usermodel.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.io.FileInputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @Author: HuangC
 * @Description:
 * @Date: 2018/6/14 13:52
 */
public class ExcelUtil1 {
    /**
     * 将用户的信息导入到excel文件中去
     * @param userList 用户列表
     * @param out 输出表
     */
    public static void exportUserExcel(List<User> userList, OutputStream out) {
        try{
            //1.创建工作簿
            HSSFWorkbook workbook = new HSSFWorkbook();
            //1.1创建合并单元格对象
            CellRangeAddress callRangeAddress = new CellRangeAddress(0,0,0,4);//起始行,结束行,起始列,结束列
            //1.2头标题样式
            HSSFCellStyle headStyle = createCellStyle(workbook,(short)16);
            //1.3列标题样式
            HSSFCellStyle colStyle = createCellStyle(workbook,(short)13);
            //2.创建工作表
            HSSFSheet sheet = workbook.createSheet("用户列表");
            //2.1加载合并单元格对象
            sheet.addMergedRegion(callRangeAddress);
            //设置默认列宽
            sheet.setDefaultColumnWidth(25);
            //3.创建行
            //3.1创建头标题行;并且设置头标题
            HSSFRow row = sheet.createRow(0);
            HSSFCell cell = row.createCell(0);

            //加载单元格样式
            cell.setCellStyle(headStyle);
            cell.setCellValue("用户列表");

            //3.2创建列标题;并且设置列标题
            HSSFRow row2 = sheet.createRow(1);
            String[] titles = {"用户名","账号","所属部门","性别","电子邮箱"};
            for(int i=0;i<titles.length;i++) {
                HSSFCell cell2 = row2.createCell(i);
                //加载单元格样式
                cell2.setCellStyle(colStyle);
                cell2.setCellValue(titles[i]);
            }


            //4.操作单元格;将用户列表写入excel
            if(userList != null) {
                for(int j=0;j<userList.size();j++) {
                    //创建数据行,前面有两行,头标题行和列标题行
                    HSSFRow row3 = sheet.createRow(j+2);
                    HSSFCell cell1 = row3.createCell(0);
                    cell1.setCellValue(userList.get(j).getName());
                    HSSFCell cell2 = row3.createCell(1);
                    cell2.setCellValue(userList.get(j).getAccount());
                    HSSFCell cell3 = row3.createCell(2);
                    cell3.setCellValue(userList.get(j).getDept());
                    HSSFCell cell4 = row3.createCell(3);
                    cell4.setCellValue(userList.get(j).isGender()?"男":"女");
                    HSSFCell cell5 = row3.createCell(4);
                    cell5.setCellValue(userList.get(j).getEmail());
                }
            }
            //5.输出
            workbook.write(out);
            //6.關閉
            workbook.close();
            out.close();
        }catch(Exception e) {
            e.printStackTrace();
        }
    }

    /**
     *
     * @param workbook
     * @param fontsize
     * @return 单元格样式
     */
    private static HSSFCellStyle createCellStyle(HSSFWorkbook workbook, short fontsize) {
        // TODO Auto-generated method stub
        HSSFCellStyle style = workbook.createCellStyle();
        style.setAlignment(HSSFCellStyle.ALIGN_CENTER);//水平居中
        style.setVerticalAlignment(HSSFCellStyle.VERTICAL_CENTER);//垂直居中
        //创建字体
        HSSFFont font = workbook.createFont();
        font.setBoldweight(HSSFFont.BOLDWEIGHT_BOLD);
        font.setFontHeightInPoints(fontsize);
        //加载字体
        style.setFont(font);
        return style;
    }



    public static void importExcel(String filePath ) {
        // TODO Auto-generated method stub
        //1.创建输入流
        try {
            FileInputStream inputStream = new FileInputStream(filePath);
            //boolean is03Excel = excelFileName.matches("^.+\\.(?i)(xls)$");
            //1.读取工作簿
            //Workbook workbook = is03Excel ? new HSSFWorkbook(inputStream) : new XSSFWorkbook(inputStream);
            Workbook workbook = new HSSFWorkbook(inputStream);
            //2.读取工作表
            Sheet sheet = workbook.getSheetAt(0);
            //3.读取行
            //判断行数大于二,是因为数据从第三行开始插入
            if (sheet.getPhysicalNumberOfRows() > 2) {
                List<User> list = new ArrayList<User>();
                User user = null;
                //跳过前两行
                for (int k = 2; k < sheet.getPhysicalNumberOfRows(); k++) {
                    //读取单元格
                    Row row = sheet.getRow(k);
                    user = new User();
                    //用户名
                    Cell cell0 = row.getCell(0);
                    user.setName(cell0.getStringCellValue());
                    //账号
                    Cell cell1 = row.getCell(1);
                    user.setAccount(cell1.getStringCellValue());
                    //所属部门
                    Cell cell2 = row.getCell(2);
                    user.setDept(cell2.getStringCellValue());
                    //设置性别
                    Cell cell3 = row.getCell(3);
                    boolean gender = "男".equals(cell3.getStringCellValue())  ? false : true;
                    user.setGender(gender);
                    //设置电子邮箱
                    Cell cell4 = row.getCell(4);
                    user.setEmail(cell4.getStringCellValue());

                    list.add(user);

                }
                System.out.println(list.toString());
            }
            workbook.close();
            inputStream.close();
        } catch (Exception e) {
            // TODO Auto-generated catch block
            e.printStackTrace();
        }
    }
}

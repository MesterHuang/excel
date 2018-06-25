package excelTest;


import excelUtil.PayCost;
import excelUtil.PayCostDetail;
import excelUtil.PayCostItem;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;

import static excelUtil.ExcelUtil2.exportUserExcel;
import static excelUtil.ExcelUtil1.importExcel;

/**
 * @Author: HuangC
 * @Description:
 * @Date: 2018/6/14 16:40
 */
public class Test2 {

    public static void main(String[] args) {
        List<PayCostDetail> payCostDetailList = new ArrayList<PayCostDetail>();
        PayCostDetail payCostDetail = new PayCostDetail();
        payCostDetail.setContent("書本費");
        payCostDetail.setAmount("200");
        payCostDetail.setCreateTime("2018-02-15");

        PayCostDetail payCostDetail1 = new PayCostDetail();
        payCostDetail1.setContent("生活費");
        payCostDetail1.setAmount("200");
        payCostDetail1.setCreateTime("2018-02-15");

        payCostDetailList.add(payCostDetail);
        payCostDetailList.add(payCostDetail1);


        List<PayCostItem> payCostItemList = new ArrayList<PayCostItem>();
        PayCostItem payCostItem = new PayCostItem();
        payCostItem.setName("學費1");
        payCostItem.setCreateTime("2018-02-15");
        payCostItem.setPayCostDetails(payCostDetailList);

        PayCostItem payCostItem1 = new PayCostItem();
        payCostItem1.setName("學費2");
        payCostItem1.setCreateTime("2018-02-15");
        payCostItem1.setPayCostDetails(payCostDetailList);

        payCostItemList.add(payCostItem);
        payCostItemList.add(payCostItem1);

        List<PayCost> payCostList = new ArrayList<PayCost>();
        PayCost payCost = new PayCost();
        payCost.setType("教育");
        payCost.setAmount("2000");
        payCost.setName("2018年學費");
        payCost.setEndTime("2018-06-10");
        payCost.setStudentCode("007");
        payCost.setStudentName("陈小伟");
        payCost.setSchoolName("香港朝阳小学");
        payCost.setStatus(1);
        payCost.setCreateTime("2018-02-15");
        payCost.setPayCostItems(payCostItemList);

        payCostList.add(payCost);

        OutputStream os = null;
        try {
            os = new FileOutputStream("E:"+ File.separator +"264.xls");
        } catch (FileNotFoundException e) {
            e.printStackTrace();
        }
        exportUserExcel(payCostList,os);
        //importExcel("E:"+ File.separator+"254.xls");
    }
}

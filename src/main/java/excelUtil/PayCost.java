package excelUtil;

import lombok.Data;

import java.util.List;

/**
 * @Author: HuangC
 * @Description:
 * @Date: 2018/6/14 16:41
 */
@Data
public class PayCost {

    private String type;
    private String name;
    private String amount;
    private String endTime;
    private String studentCode;
    private String studentName;
    private String schoolName;
    private Integer status;
    private String createTime;

    private List<PayCostItem> payCostItems;
}

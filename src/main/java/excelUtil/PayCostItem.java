package excelUtil;

import lombok.Data;

import java.util.List;

/**
 * @Author: HuangC
 * @Description:
 * @Date: 2018/6/14 16:45
 */
@Data
public class PayCostItem {

    private String name;
    private String createTime;
    private List<PayCostDetail> payCostDetails;
}

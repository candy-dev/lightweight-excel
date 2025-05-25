package csc.com.demo;

import csc.com.annotation.ExcelField;
import csc.com.annotation.ExcelSheet;
import lombok.Data;

import java.math.BigDecimal;
import java.util.Date;

@Data
@ExcelSheet(name = "交易记录", moneyUnit = "元")
public   class TransactionEntity {
    @ExcelField(title = "序号")
    private Integer index;
    @ExcelField(title = "交易物品")
    private String name;
    @ExcelField(title = "交易时间")
    private Date date;
    @ExcelField(title = "交易金额", sum = true)
    private BigDecimal money;

}
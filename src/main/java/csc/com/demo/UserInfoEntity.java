package csc.com.demo;

import csc.com.annotation.ExcelColumn;
import csc.com.annotation.ExcelField;
import csc.com.annotation.ExcelSheet;
import csc.com.converter.DateConverter;
import lombok.Data;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.Date;

/**
 * 使用示例
 *
 * @author 君心今何在
 */
@Data
@ExcelSheet(name = "用户信息报表", moneyUnit = "元")
public class UserInfoEntity {
    @ExcelColumn(value = "序号")
    @ExcelField(title = "序号")
    private Integer index;

    @ExcelColumn(value = "姓名" )
    @ExcelField(title = "姓名")
    private String name;

    @ExcelColumn(value = "年龄")
    @ExcelField(title = "年龄")
    private Integer age;

    @ExcelColumn(value = "性别")
    @ExcelField(title = "性别")
    private Integer sex;

    @ExcelColumn(value = "身高")
    @ExcelField(title = "身高")
    private Integer height;

    @ExcelColumn(value = "体重")
    @ExcelField(title = "体重")
    private Integer weight;

    @ExcelColumn(value = "邮箱")
    @ExcelField(title = "邮箱")
    private String email;

    @ExcelColumn(value = "入职日期", converter = DateConverter.class)
    @ExcelField(title = "入职日期", dateFormat = "yyyy-MM-dd")
    private LocalDate joinDate;

    @ExcelColumn(value = "入职公司")
    @ExcelField(title = "入职公司")
    private String companyName;

    @ExcelColumn(value = "薪资")
    @ExcelField(title = "薪资", sum = true)
    private BigDecimal salary;

    @ExcelColumn(value = "存款")
    @ExcelField(title = "存款", sum = true)
    private BigDecimal sumMoney;

    @ExcelColumn(value = "是否在职", defaultValue = "true")
    @ExcelField(title = "是否在职")
    private Boolean active;

    @ExcelColumn(value = "联系电话")
    @ExcelField(title = "联系电话")
    private String phone;

    @ExcelColumn(value = "注册时间", converter = DateConverter.class)
    @ExcelField(title = "注册时间", dateFormat = "yyyy-MM-dd")
    private LocalDate registeredDate;

    @ExcelColumn(value = "备注")
    @ExcelField(title = "备注" )
    private String bz;
    @ExcelColumn(value = "备注2")
    @ExcelField(title = "备注2" )
    private String bz2;
    @ExcelColumn(value = "备注3")
    @ExcelField(title = "备注3" )
    private String bz3;
    @ExcelColumn(value = "备注4")
    @ExcelField(title = "备注4" )
    private String bz4;
    @ExcelColumn(value = "备注5")
    @ExcelField(title = "备注5" )
    private String bz5;
    @ExcelColumn(value = "备注6")
    @ExcelField(title = "备注6" )
    private String bz6;
    @ExcelColumn(value = "备注7")
    @ExcelField(title = "备注7" )
    private String bz7;
    @ExcelColumn(value = "备注8")
    @ExcelField(title = "备注8" )
    private String bz8;
    @ExcelColumn(value = "备注9")
    @ExcelField(title = "备注9" )
    private String bz9;
    @ExcelColumn(value = "备注10")
    @ExcelField(title = "备注10" )
    private String bz10;
    @ExcelColumn(value = "备注11")
    @ExcelField(title = "备注12" )
    private String bz11;

}
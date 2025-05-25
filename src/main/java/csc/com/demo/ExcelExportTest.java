package csc.com.demo;

import csc.com.util.ExcelExportUtil;
import org.apache.poi.ss.formula.functions.T;

import java.io.FileOutputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

/**
 * 导出测试类
 * @author 刺猬
 */
public class ExcelExportTest {

    public static void main(String[] args) {
        long start = System.currentTimeMillis();

            // 导出到文件

                String fileName = "export_data.xlsx";
               //多sheet导出 准备多个Sheet的数据
                List<ExcelExportUtil.SheetData<T>> exportDataList = new ArrayList<>();
                exportDataList.add(createSalesSheetData());
                exportDataList.add(createUerInfoSheetData());
                ExcelExportUtil.exportMultiSheet(fileName, exportDataList);

                //单sheet导出
//                ExcelExportUtil.export(fileName, buildUserInfoData(), UserInfoEntity.class);

                System.out.println("Excel导出成功：export_data.xlsx");


        long endTime = System.currentTimeMillis();

        System.out.println("耗时：" + (endTime - start) / 1000 + "s");


    }

    /**
     * 创建销售数据sheet
     */
    private static ExcelExportUtil.SheetData  createSalesSheetData() {

        String[] transaction =  {"可口可乐","哇哈哈","沙漠西瓜","冰红茶","红酒","啤酒","塔斯丁","杜仲酒","张家界莓茶","混合坚果"};

        List<TransactionEntity> dataList = new ArrayList<>();
        Date now = new Date();
        for (int i = 0 ;i <= 5; i++){

            TransactionEntity deom = new TransactionEntity();

            int index = i % 10;

            deom.setIndex(i);
            deom.setName(transaction[index]);
            deom.setMoney(new BigDecimal("1000"));
            deom.setDate(now);
            dataList.add(deom);
        }

        // 创建ExportData对象
        return new ExcelExportUtil.SheetData(dataList, TransactionEntity.class);
    }

    /**
     * 创建用户信息数据sheet
     */
    private static  ExcelExportUtil.SheetData  createUerInfoSheetData() {
        // 创建ExportData对象
        return new ExcelExportUtil.SheetData (buildUserInfoData(), UserInfoEntity.class);
    }

    /**
     * 创建用户信息数据
     */
    private static  List<UserInfoEntity>  buildUserInfoData() {
        // 设置数据
        List<UserInfoEntity> dataList = new ArrayList<>();

        LocalDate now = LocalDate.now();

        for (long i = 0 ;i <= 5; i++){
            UserInfoEntity demo = new UserInfoEntity();
            demo.setIndex((int) i);
            demo.setName("张三"+i);
            demo.setAge(100);
            demo.setSex(1);
            demo.setHeight(182);
            demo.setWeight(65);
            demo.setEmail("aa@qq.com");
            demo.setJoinDate(now);
            demo.setCompanyName("刺猬集团");
            demo.setSalary(new BigDecimal(1000));
            demo.setSumMoney(new BigDecimal(999999999));
            demo.setActive(true);
            demo.setPhone("13999999999");
            demo.setRegisteredDate(now);
            demo.setBz("这里是备注，是备注呀 this is remark");
            demo.setBz2("这里是备注，是备注呀 this is remark");
            demo.setBz3("这里是备注，是备注呀 this is remark");
            demo.setBz4("这里是备注，是备注呀 this is remark");
            demo.setBz5("这里是备注，是备注呀 this is remark");
            demo.setBz6("这里是备注，是备注呀 this is remark");
            demo.setBz7("这里是备注，是备注呀 this is remark");
            demo.setBz8("这里是备注，是备注呀 this is remark");
            demo.setBz9("这里是备注，是备注呀 this is remark");
            demo.setBz10("这里是备注，是备注呀 this is remark");
            demo.setBz11("这里是备注，是备注呀 this is remark");
            dataList.add(demo);
        }
        // 创建ExportData对象
        return dataList;
    }





}

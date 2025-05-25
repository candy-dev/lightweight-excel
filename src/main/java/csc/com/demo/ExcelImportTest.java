package csc.com.demo;

import csc.com.common.ImportResult;
import csc.com.util.ExcelImportUtil;

import java.io.FileInputStream;
import java.util.Arrays;
import java.util.List;
import java.util.Map;

/**
 * 导出测试类
 *
 * @author 刺猬
 */
public class ExcelImportTest {
    public static void main(String[] args) {
        try {
            // 测试导入第一个Sheet
            String fileName = "export_data.xlsx";
            ImportResult<UserInfoEntity> result = ExcelImportUtil.importFirstSheet(fileName, UserInfoEntity.class);

            // 打印成功的数据
            result.getSuccessList().forEach(System.out::println);
            // 打印错误信息
            result.getErrorList().forEach(System.out::println);

            System.out.println("Sheet: " + result.getSheetName());
            System.out.println("总记录数: " + result.getTotalCount());
            System.out.println("成功记录数: " + result.getSuccessCount());
            System.out.println("失败记录数: " + result.getErrorCount());


        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}
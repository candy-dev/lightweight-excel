package csc.com.demo;

import csc.com.common.ImportResult;
import csc.com.common.MultiSheetImportResult;
import csc.com.util.ExcelImportUtil;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.*;

/**
 * 导出测试类
 * @author 刺猬
 */
public class ExcelImportMultiTest {
    public static void main(String[] args) {
        System.out.println("\n=== 多Sheet导入示例 ===");

        // 配置多个Sheet的类映射
        Map<String, Class<?>> sheetClassMap = new HashMap<>();
        sheetClassMap.put("交易记录", TransactionEntity.class);
        sheetClassMap.put("用户信息报表", UserInfoEntity.class);

        String fileName = "export_data.xlsx";

        MultiSheetImportResult result = ExcelImportUtil.importMultiSheet(fileName,   sheetClassMap);

        if (result.isSuccess()) {
            System.out.println("多Sheet导入成功！");
        } else {
            System.out.println("多Sheet导入失败: " + result.getMessage());
            for (String error : result.getGlobalErrors()) {
                System.out.println("全局错误: " + error);
            }
        }

        // 处理每个Sheet的结果
        for (Map.Entry<String, ImportResult<?>> entry : result.getSheetResults().entrySet()) {
            String sheetName = entry.getKey();
            ImportResult<?> sheetResult = entry.getValue();

            System.out.println("\n--- Sheet: " + sheetName + " ---");
            System.out.println("状态: " + (sheetResult.isSuccess() ? "成功" : "失败"));
            System.out.println("数量: " + (sheetResult.isSuccess() ? sheetResult.getSuccessCount() : sheetResult.getErrorCount()));
            sheetResult.getSuccessList().forEach(System.out::println);

        }

    }
}
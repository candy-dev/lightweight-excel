package csc.com.common;

import lombok.Data;

import java.util.List;
import java.util.Map;

/**
 * 多Sheet导入结果
 * @author 刺猬
 */
@Data
public class MultiSheetImportResult {
    private boolean success;
    private String message;
    private Map<String, ImportResult<?>> sheetResults;
    private List<String> globalErrors;


}
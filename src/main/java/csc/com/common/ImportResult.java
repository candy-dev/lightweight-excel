package csc.com.common;

import java.util.ArrayList;
import java.util.List;

/**
 * 导入结果
 * @param <T>
 * @author 刺猬
 */
public class ImportResult<T> {
    private List<T> successList = new ArrayList<>();
    private List<ImportError> errorList = new ArrayList<>();
    private int totalCount;
    private String sheetName;
    
    // 构造函数、getter、setter
    public ImportResult() {}
    
    public ImportResult(String sheetName) {
        this.sheetName = sheetName;
    }
    
    public List<T> getSuccessList() { return successList; }
    public void setSuccessList(List<T> successList) { this.successList = successList; }
    
    public List<ImportError> getErrorList() { return errorList; }
    public void setErrorList(List<ImportError> errorList) { this.errorList = errorList; }
    
    public int getTotalCount() { return totalCount; }
    public void setTotalCount(int totalCount) { this.totalCount = totalCount; }
    
    public String getSheetName() { return sheetName; }
    public void setSheetName(String sheetName) { this.sheetName = sheetName; }
    
    public boolean isSuccess() {
        return errorList.isEmpty();
    }
    
    public int getSuccessCount() {
        return successList.size();
    }
    
    public int getErrorCount() {
        return errorList.size();
    }
}

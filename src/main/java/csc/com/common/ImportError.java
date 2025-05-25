package csc.com.common;

/**
 * 错误信息
 * @author 刺猬
 */
public class ImportError {
    private int rowNum;
    private String columnName;
    private String cellValue;
    private String errorMessage;
    
    public ImportError(int rowNum, String columnName, String cellValue, String errorMessage) {
        this.rowNum = rowNum;
        this.columnName = columnName;
        this.cellValue = cellValue;
        this.errorMessage = errorMessage;
    }
    
    // getter、setter
    public int getRowNum() { return rowNum; }
    public void setRowNum(int rowNum) { this.rowNum = rowNum; }
    
    public String getColumnName() { return columnName; }
    public void setColumnName(String columnName) { this.columnName = columnName; }
    
    public String getCellValue() { return cellValue; }
    public void setCellValue(String cellValue) { this.cellValue = cellValue; }
    
    public String getErrorMessage() { return errorMessage; }
    public void setErrorMessage(String errorMessage) { this.errorMessage = errorMessage; }
    
    @Override
    public String toString() {
        return String.format("第%d行[%s]列值[%s]错误: %s", rowNum, columnName, cellValue, errorMessage);
    }
}
package csc.com.util;

import csc.com.annotation.ExcelColumn;
import csc.com.annotation.ExcelSheet;
import csc.com.common.ImportError;
import csc.com.common.ImportResult;
import csc.com.common.MultiSheetImportResult;
import csc.com.converter.DataConverter;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.Field;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * 通用Excel导入工具
 * @author 刺猬
 */
public class ExcelImportUtil {


    /**
     * 导入第一个Sheet
     */
    public static <T> ImportResult<T> importFirstSheet(String fileName, Class<T> clazz) {
        try (FileInputStream inputStream = new FileInputStream(fileName); Workbook workbook = createWorkbook(inputStream)) {
            if (workbook.getNumberOfSheets() > 0) {
                Sheet firstSheet = workbook.getSheetAt(0);
                return parseSheet(firstSheet, clazz);
            } else {
                ImportResult<T> result = new ImportResult<>("未找到Sheet");
                result.getErrorList().add(new ImportError(0, "", "", "Excel文件中未找到任何Sheet"));
                return result;
            }
        } catch (Exception e) {
            ImportResult<T> result = new ImportResult<>("解析异常");
            result.getErrorList().add(new ImportError(0, "", "", "解析Excel文件异常: " + e.getMessage()));
            return result;
        }
    }


    /**
     * 导入指定Sheet
     */
    public static <T> ImportResult<T> importSheet(String fileName, String sheetName, Class<T> clazz) {
        Map<String, ImportResult<T>> results = importMultiBySheet(fileName, sheetName, clazz);
        return results.getOrDefault(sheetName, new ImportResult<>(sheetName));
    }

    /**
     * 导入指定
     */
    public static <T> Map<String, ImportResult<T>> importMultiBySheet(String fileName, String sheetName, Class<T> clazz) {
        Map<String, ImportResult<T>> results = new LinkedHashMap<>();

        try (FileInputStream inputStream = new FileInputStream(fileName); Workbook workbook = createWorkbook(inputStream)) {
            if (sheetName == null ) {
                //没有指定则导入第一个
                importFirstSheet(fileName,clazz);
            } else {
                // 导入指定Sheet

                    try {
                        Sheet sheet = workbook.getSheet(sheetName);
                        if (sheet != null) {
                            ImportResult<T> result = parseSheet(sheet, clazz);
                            results.put(sheetName, result);
                        } else {
                            ImportResult<T> errorResult = new ImportResult<>(sheetName);
                            errorResult.getErrorList().add(new ImportError(0, "", "", "未找到名为'" + sheetName + "'的Sheet"));
                            results.put(sheetName, errorResult);
                        }
                    } catch (Exception e) {
                        ImportResult<T> errorResult = new ImportResult<>(sheetName);
                        errorResult.getErrorList().add(new ImportError(0, "", "", "解析Sheet[" + sheetName + "]异常: " + e.getMessage()));
                        results.put(sheetName, errorResult);
                    }

            }
        } catch (Exception e) {
            ImportResult<T> errorResult = new ImportResult<>("全局异常");
            errorResult.getErrorList().add(new ImportError(0, "", "", "解析Excel文件异常: " + e.getMessage()));
            results.put("ERROR", errorResult);
        }

        return results;
    }

    /**
     * 多Sheet导入
     */
    public static MultiSheetImportResult importMultiSheet(String fileName, Map<String, Class<?>> sheetClassMap) {
        MultiSheetImportResult result = new MultiSheetImportResult();
        Map<String, ImportResult<?>> sheetResults = new HashMap<>();
        List<String> globalErrors = new ArrayList<>();
        boolean allSuccess = true;

        try(FileInputStream inputStream = new FileInputStream(fileName); Workbook workbook = createWorkbook(inputStream)) {

            for (Map.Entry<String, Class<?>> entry : sheetClassMap.entrySet()) {
                String sheetName = entry.getKey();
                Class<?> clazz = entry.getValue();

                try {
                    Sheet sheet = workbook.getSheet(sheetName);
                    if (sheet == null) {
                        // 尝试通过注解获取Sheet
                        ExcelSheet sheetConfig = clazz.getAnnotation(ExcelSheet.class);
                        if (sheetConfig != null && sheetConfig.index() >= 0) {
                            sheet = workbook.getSheetAt(sheetConfig.index());
                        }
                    }

                    if (sheet != null) {
                        ImportResult<?> sheetResult = parseSheet(sheet, clazz);
                        sheetResults.put(sheetName, sheetResult);
                        if (!sheetResult.isSuccess()) {
                            allSuccess = false;
                        }
                    } else {
                        globalErrors.add("未找到Sheet: " + sheetName);
                        allSuccess = false;
                    }

                } catch (Exception e) {
                    globalErrors.add("解析Sheet[" + sheetName + "]失败: " + e.getMessage());
                    allSuccess = false;
                }
            }

        } catch (Exception e) {
            globalErrors.add("读取Excel文件失败: " + e.getMessage());
            allSuccess = false;
        }

        result.setSuccess(allSuccess);
        result.setMessage(allSuccess ? "导入成功" : "部分导入失败");
        result.setSheetResults(sheetResults);
        result.setGlobalErrors(globalErrors);

        return result;
    }

    /**
     * 创建Workbook对象
     */
    private static Workbook createWorkbook(InputStream inputStream) throws IOException {
        // 支持.xls和.xlsx格式
        try {
            return new XSSFWorkbook(inputStream);
        } catch (Exception e) {
            // 如果XLSX格式失败，尝试XLS格式
            try {
                inputStream.reset();
                return new HSSFWorkbook(inputStream);
            } catch (Exception ex) {
                throw new IOException("不支持的Excel文件格式", ex);
            }
        }
    }

    /**
     * 解析单个Sheet
     */
    private static <T> ImportResult<T> parseSheet(Sheet sheet, Class<T> clazz) {
        ImportResult<T> result = new ImportResult<>(sheet.getSheetName());

        try {
            int lastRowNum = sheet.getLastRowNum();

            if (lastRowNum < 2) {
                result.getErrorList().add(new ImportError(0, "", "", "Excel数据行数不足，至少需要3行（标题在第3行）"));
                return result;
            }

            // 第3行为标题行（索引为2）
            Row headerRow = sheet.getRow(2);
            if (headerRow == null) {
                result.getErrorList().add(new ImportError(3, "", "", "第3行标题行为空"));
                return result;
            }

            Map<String, Integer> headerMap = parseHeader(headerRow);

            // 获取字段映射
            Map<String, FieldInfo> fieldMap = parseFieldAnnotations(clazz);

            // 验证必需的列是否存在
            validateRequiredColumns(fieldMap, headerMap, result);
            if (!result.getErrorList().isEmpty()) {
                return result;
            }

            // 从第4行开始解析数据（索引从3开始）
            for (int i = 3; i <= lastRowNum; i++) {
                Row dataRow = sheet.getRow(i);
                int rowNum = i + 1; // Excel行号从1开始

                if (dataRow == null) {
                    continue; // 跳过空行
                }

//                // 检查是否为合计行
//                if (hasNumericColumns && isSummaryRow(dataRow, columnMap, fieldMap)) {
//                    break; // 遇到合计行，停止解析
//                }

                try {
                    T obj = parseRowData(dataRow, headerMap, fieldMap, clazz, rowNum, result);
                    if (obj != null) {
                        result.getSuccessList().add(obj);
                    }
                } catch (Exception e) {
                    result.getErrorList().add(new ImportError(rowNum, "", "", "解析行数据异常: " + e.getMessage()));
                }
            }

            result.setTotalCount(lastRowNum - 2); // 减去前3行（0,1,2）

        } catch (Exception e) {
            result.getErrorList().add(new ImportError(0, "", "", "解析Sheet异常: " + e.getMessage()));
        }

        return result;
    }


    /**
     * 解析标题行
     */
    private static Map<String, Integer> parseHeader(Row headerRow) {
        Map<String, Integer> headerMap = new HashMap<>();

        for (Cell cell : headerRow) {
            String headerName = getCellValue(cell);
            if (StringUtils.isNotBlank(headerName)) {
                headerMap.put(headerName.trim(), cell.getColumnIndex());
            }
        }

        return headerMap;
    }

    /**
     * 获取单元格值
     */
    private static String getCellValue(Cell cell) {
        if (cell == null) {
            return "";
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // 日期类型，格式化为字符串
                    Date date = cell.getDateCellValue();
                    return DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")
                            .format(date.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime());
                } else {
                    double numericValue = cell.getNumericCellValue();
                    if (numericValue == (long) numericValue) {
                        return String.valueOf((long) numericValue);
                    } else {
                        return String.valueOf(numericValue);
                    }
                }
            case BOOLEAN:
                return String.valueOf(cell.getBooleanCellValue());
            case FORMULA:
                // 对于公式，获取计算结果
                try {
                    FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                    CellValue cellValue = evaluator.evaluate(cell);
                    switch (cellValue.getCellType()) {
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                Date date = DateUtil.getJavaDate(cellValue.getNumberValue());
                                return DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")
                                        .format(date.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime());
                            } else {
                                double numValue = cellValue.getNumberValue();
                                if (numValue == (long) numValue) {
                                    return String.valueOf((long) numValue);
                                } else {
                                    return String.valueOf(numValue);
                                }
                            }
                        case STRING:
                            return cellValue.getStringValue();
                        case BOOLEAN:
                            return String.valueOf(cellValue.getBooleanValue());
                        default:
                            return "";
                    }
                } catch (Exception e) {
                    return cell.getCellFormula();
                }
            default:
                return "";
        }
    }

    /**
     * 获取单元格原始值（用于类型转换）
     */
    private static Object getCellRawValue(Cell cell) {
        if (cell == null) {
            return null;
        }

        switch (cell.getCellType()) {
            case STRING:
                return cell.getStringCellValue();
            case NUMERIC:
                if (DateUtil.isCellDateFormatted(cell)) {
                    // 返回Date对象，让转换器处理
                    return cell.getDateCellValue();
                } else {
                    // 返回数值，可能是日期序列号
                    return cell.getNumericCellValue();
                }
            case BOOLEAN:
                return cell.getBooleanCellValue();
            case FORMULA:
                // 对于公式，获取计算后的值
                try {
                    FormulaEvaluator evaluator = cell.getSheet().getWorkbook().getCreationHelper().createFormulaEvaluator();
                    CellValue cellValue = evaluator.evaluate(cell);
                    switch (cellValue.getCellType()) {
                        case NUMERIC:
                            if (DateUtil.isCellDateFormatted(cell)) {
                                return DateUtil.getJavaDate(cellValue.getNumberValue());
                            } else {
                                return cellValue.getNumberValue();
                            }
                        case STRING:
                            return cellValue.getStringValue();
                        case BOOLEAN:
                            return cellValue.getBooleanValue();
                        default:
                            return null;
                    }
                } catch (Exception e) {
                    // 如果公示计算失败，返回公式字符串
                    return cell.getCellFormula();
                }
            default:
                return null;
        }
    }

    /**
     * 解析字段注解
     */
    private static Map<String, FieldInfo> parseFieldAnnotations(Class<?> clazz) {
        Map<String, FieldInfo> fieldMap = new HashMap<>();

        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
            if (annotation != null) {
                field.setAccessible(true);
                FieldInfo fieldInfo = new FieldInfo(field, annotation);
                fieldMap.put(annotation.value(), fieldInfo);
            }
        }

        return fieldMap;
    }

    /**
     * 验证必需的列
     */
    private static void validateRequiredColumns(Map<String, FieldInfo> fieldMap, Map<String, Integer> headerMap, ImportResult<?> result) {
        for (Map.Entry<String, FieldInfo> entry : fieldMap.entrySet()) {
            String columnName = entry.getKey();
            FieldInfo fieldInfo = entry.getValue();

            if (fieldInfo.getAnnotation().required() && !headerMap.containsKey(columnName)) {
                result.getErrorList().add(new ImportError(3, columnName, "", "缺少必需的列: " + columnName));
            }
        }
    }

    /**
     * 解析行数据
     */
    private static <T> T parseRowData(Row dataRow, Map<String, Integer> headerMap,
                                      Map<String, FieldInfo> fieldMap, Class<T> clazz,
                                      int rowNum, ImportResult<T> result) throws Exception {

        T obj = clazz.getDeclaredConstructor().newInstance();
        boolean hasData = false;

        for (Map.Entry<String, FieldInfo> entry : fieldMap.entrySet()) {
            String columnName = entry.getKey();
            FieldInfo fieldInfo = entry.getValue();
            Field field = fieldInfo.getField();
            ExcelColumn annotation = fieldInfo.getAnnotation();

            Integer columnIndex = headerMap.get(columnName);
            if (columnIndex == null) {
                continue; // 列不存在，跳过
            }

            String cellDisplayValue = "";
            try {
                Cell cell = dataRow.getCell(columnIndex);
                Object cellValue = getCellRawValue(cell);
                cellDisplayValue = getCellValue(cell);

                // 处理默认值
                if ((cellValue == null || StringUtils.isBlank(cellDisplayValue)) && StringUtils.isNotBlank(annotation.defaultValue())) {
                    cellValue = annotation.defaultValue();
                    cellDisplayValue = annotation.defaultValue();
                }

                // 验证必填
                if (annotation.required() && (cellValue == null || StringUtils.isBlank(cellDisplayValue))) {
                    result.getErrorList().add(new ImportError(rowNum, columnName, cellDisplayValue, "必填字段不能为空"));
                    continue;
                }

                // 转换数据
                if (cellValue != null && StringUtils.isNotBlank(cellDisplayValue)) {
                    hasData = true;
                    DataConverter converter = annotation.converter().getDeclaredConstructor().newInstance();
                    Object value = converter.convert(cellValue, field);
                    field.set(obj, value);
                }

            } catch (Exception e) {
                result.getErrorList().add(new ImportError(rowNum, columnName, cellDisplayValue,
                        "数据转换失败: " + e.getMessage()));
            }
        }

        return hasData ? obj : null; // 如果整行都没有数据，返回null
    }

    // 字段信息内部类
    private static class FieldInfo {
        private Field field;
        private ExcelColumn annotation;

        public FieldInfo(Field field, ExcelColumn annotation) {
            this.field = field;
            this.annotation = annotation;
        }

        public Field getField() { return field; }
        public ExcelColumn getAnnotation() { return annotation; }
    }
}
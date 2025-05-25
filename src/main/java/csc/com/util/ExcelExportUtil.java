package csc.com.util;

import csc.com.annotation.ExcelField;
import csc.com.annotation.ExcelSheet;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.*;

/**
 * 基于注解的Apache POI通用导出工具
 * 基于POI 4.1.2版本
 * 支持多sheet导出
 * 自动处理第一行为标题，第二行第一列为导出时间，最后一列为金额单位
 * 自动为数值列添加表尾合计
 *
 * @author 刺猬
 */
public class ExcelExportUtil {

    // 日期格式化
    private static final DateTimeFormatter DEFAULT_DATE_FORMATTER = DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss");

    /**
     * 导出单个对象列表到Excel
     *
     * @param fileName 导出的文件名
     * @param dataList     数据列表
     * @param clazz        数据类型Class
     * @throws Exception 异常
     */
    public static <T> void export(String fileName, List<T> dataList, Class<T> clazz)   {
        List<SheetData<T>> sheetDataList = new ArrayList<>();
        sheetDataList.add(new SheetData<>(dataList, clazz));
        exportMultiSheet(fileName, sheetDataList);
    }

    /**
     * 导出多个Sheet到Excel
     *
     * @param fileName  导出的文件名
     * @param sheetDataList Sheet数据列表
     * @throws Exception 异常
     */
    public static <T> void exportMultiSheet(String fileName, List<SheetData<T>> sheetDataList) {
        if (sheetDataList == null || sheetDataList.isEmpty()) {
            throw new IllegalArgumentException("导出数据不能为空");
        }

        // 使用SXSSF模式创建工作簿，适合大数据量导出, 设置内存中保留100行，其余写入临时文件
        try (SXSSFWorkbook workbook = new SXSSFWorkbook(100);
             FileOutputStream outputStream = new FileOutputStream(fileName)) {
            // 创建常用样式
            Map<String, CellStyle> styles = createStyles(workbook);

            // 为每个导出数据创建工作表
            for (SheetData<?> sheetData : sheetDataList) {
                createSheet(workbook, sheetData, styles);
            }

            // 写入输出流
            workbook.write(outputStream);

            // 清理临时文件
            workbook.dispose();
        } catch (FileNotFoundException e) {
            throw new RuntimeException(e);
        } catch (IOException e) {
            throw new RuntimeException(e);
        } catch (Exception e) {
            throw new RuntimeException(e);
        }
    }

    /**
     * 创建通用样式
     *
     * @param workbook 工作簿
     * @return 样式映射
     */
    private static Map<String, CellStyle> createStyles(Workbook workbook) {
        Map<String, CellStyle> styles = new HashMap<>();
        //标题样式
        CellStyle titleStyle = workbook.createCellStyle();
        Font titleFont = workbook.createFont();
        titleFont.setBold(true);
        titleFont.setFontHeightInPoints((short) 28);
        titleStyle.setFont(titleFont);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);
        titleStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        titleStyle.setFillForegroundColor(IndexedColors.WHITE1.getIndex());
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        styles.put("title", titleStyle);


        // 表头样式
        CellStyle headerStyle = workbook.createCellStyle();
        Font headerFont = workbook.createFont();
        headerFont.setBold(true);
        headerFont.setFontHeightInPoints((short) 12);
        headerStyle.setFont(headerFont);
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        styles.put("header", headerStyle);

        // 信息行样式
        CellStyle infoStyle = workbook.createCellStyle();
        Font infoFont = workbook.createFont();
        infoFont.setFontHeightInPoints((short) 10);
        infoStyle.setFont(infoFont);
        infoStyle.setAlignment(HorizontalAlignment.LEFT);
        styles.put("info", infoStyle);

        // 默认数据样式
        CellStyle defaultStyle = workbook.createCellStyle();
        defaultStyle.setBorderTop(BorderStyle.THIN);
        defaultStyle.setBorderRight(BorderStyle.THIN);
        defaultStyle.setBorderBottom(BorderStyle.THIN);
        defaultStyle.setBorderLeft(BorderStyle.THIN);
        // 启用自动换行
        defaultStyle.setWrapText(true);
        styles.put("default", defaultStyle);

        // 数值样式（带两位小数）
        CellStyle numberStyle = workbook.createCellStyle();
        numberStyle.cloneStyleFrom(defaultStyle);
        numberStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));
        numberStyle.setAlignment(HorizontalAlignment.RIGHT);
        styles.put("number", numberStyle);

        // 整数样式（不带小数）
        CellStyle integerStyle = workbook.createCellStyle();
        integerStyle.cloneStyleFrom(defaultStyle);
        integerStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0"));
        integerStyle.setAlignment(HorizontalAlignment.RIGHT);
        styles.put("integer", integerStyle);

        // 日期样式
        CellStyle dateStyle = workbook.createCellStyle();
        dateStyle.cloneStyleFrom(defaultStyle);
        dateStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd"));
        styles.put("date", dateStyle);

        // 日期时间样式
        CellStyle dateTimeStyle = workbook.createCellStyle();
        dateTimeStyle.cloneStyleFrom(defaultStyle);
        dateTimeStyle.setDataFormat(workbook.createDataFormat().getFormat("yyyy-mm-dd hh:mm:ss"));
        styles.put("datetime", dateTimeStyle);

        // 合计行标题样式
        CellStyle summaryLabelStyle = workbook.createCellStyle();
        Font summaryFont = workbook.createFont();
        summaryFont.setBold(true);
        summaryLabelStyle.setFont(summaryFont);
        styles.put("summaryLabel", summaryLabelStyle);

        // 合计行数值样式
        CellStyle summaryNumberStyle = workbook.createCellStyle();
        summaryNumberStyle.setFont(summaryFont);
        summaryNumberStyle.setDataFormat(workbook.createDataFormat().getFormat("#,##0.00"));
        summaryNumberStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        summaryNumberStyle.setFillForegroundColor(IndexedColors.WHITE1.getIndex());
        styles.put("summaryNumber", summaryNumberStyle);

        return styles;
    }

    /**
     * 创建动态格式样式
     *
     * @param workbook  工作簿
     * @param baseStyle 基础样式
     * @param format    格式字符串
     * @return 单元格样式
     */
    private static CellStyle createFormatStyle(Workbook workbook, CellStyle baseStyle, String format) {
        CellStyle style = workbook.createCellStyle();
        style.cloneStyleFrom(baseStyle);
        style.setDataFormat(workbook.createDataFormat().getFormat(format));
        return style;
    }

    /**
     * 创建工作表
     *
     * @param workbook  工作簿
     * @param sheetData Sheet数据
     * @param styles    样式映射
     * @throws Exception 异常
     */
    private static <T> void createSheet(Workbook workbook, SheetData<T> sheetData, Map<String, CellStyle> styles) throws Exception {
        List<?> dataList = sheetData.getDataList();
        Class<?> clazz = sheetData.getClazz();

        if (dataList == null || dataList.isEmpty()) {
            return;
        }

        // 获取ExcelSheet注解
        String sheetName;
        String moneyUnit = "";
        if (clazz.isAnnotationPresent(ExcelSheet.class)) {
            ExcelSheet sheetAnnotation = clazz.getAnnotation(ExcelSheet.class);
            sheetName = sheetAnnotation.name();
            moneyUnit = sheetAnnotation.moneyUnit();

            // 如果未设置sheet名称，则使用类名
            if (sheetName == null || sheetName.trim().isEmpty()) {
                sheetName = clazz.getSimpleName();
            }
        } else {
            sheetName = clazz.getSimpleName();
        }

        // 创建工作表
        Sheet sheet = workbook.createSheet(sheetName);

        // 获取标记了ExcelField注解的字段
        List<FieldInfo> fieldInfoList = getExcelFields(clazz);
        if (fieldInfoList.isEmpty()) {
            throw new IllegalArgumentException("未找到标记了@ExcelField注解的字段");
        }

        // 设置列宽
        for (int i = 0; i < fieldInfoList.size(); i++) {
            FieldInfo fieldInfo = fieldInfoList.get(i);
            sheet.setColumnWidth(i, fieldInfo.getWidth() * 256); // POI中列宽单位为1/256个字符宽度
        }
        //写入title
        writeExportTableTitle(sheet, fieldInfoList.size(), sheetName, styles.get("title"));

        // 写入表头（第一行）
        writeHeaders(sheet, fieldInfoList, styles.get("header"));

        // 写入导出时间和金额单位（第二行）
        writeExportInfoRow(sheet, fieldInfoList.size(), moneyUnit, styles.get("info"));

        // 用于记录字段类型和自定义格式
        Map<Field, CellStyle> fieldStyles = new HashMap<>();
        Map<Integer, Boolean> needSumColumns = new HashMap<>();

        // 创建自定义格式样式
        for (FieldInfo fieldInfo : fieldInfoList) {
            Field field = fieldInfo.getField();
            // 处理自定义数值格式
            if (!fieldInfo.getNumberFormat().isEmpty()) {
                CellStyle customNumberStyle = createFormatStyle(workbook, styles.get("default"), fieldInfo.getNumberFormat());
                fieldStyles.put(field, customNumberStyle);
            }
            // 处理自定义日期格式
            else if (!fieldInfo.getDateFormat().isEmpty()) {
                CellStyle customDateStyle = createFormatStyle(workbook, styles.get("default"), fieldInfo.getDateFormat());
                fieldStyles.put(field, customDateStyle);
            }

            // 标记需要合计的列
            if (fieldInfo.isSum()) {
                needSumColumns.put(fieldInfo.getOrder(), true);
            }
        }

        // 用于存储合计值
        Map<Integer, BigDecimal> sumMap = new HashMap<>();

        // 写入数据（从第三行开始）
        int rowIndex = 3; // 第三行开始写数据（索引从0开始）
        for (Object data : dataList) {
            writeDataRow(sheet, rowIndex, data, fieldInfoList, styles, fieldStyles, sumMap);
            rowIndex++;
        }

        // 判断是否需要写入合计行
        if (!needSumColumns.isEmpty()) {
            writeSummaryRow(sheet, rowIndex, fieldInfoList, sumMap, styles);
        }
    }

    /**
     * 获取标记了ExcelField注解的字段信息
     *
     * @param clazz 类
     * @return 字段信息列表
     */
    private static List<FieldInfo> getExcelFields(Class<?> clazz) {
        List<FieldInfo> fieldInfoList = new ArrayList<>();

        // 获取所有字段（包括父类字段）
        List<Field> allFields = getAllFields(clazz);

        for (Field field : allFields) {
            ExcelField annotation = field.getAnnotation(ExcelField.class);
            if (annotation != null && !annotation.ignore()) {
                field.setAccessible(true);
                FieldInfo fieldInfo = new FieldInfo(
                        field,
                        annotation.title(),
                        annotation.order(),
                        annotation.dateFormat(),
                        annotation.numberFormat(),
                        annotation.sum(),
                        annotation.width()
                );
                fieldInfoList.add(fieldInfo);
            }
        }

        // 按order排序
        fieldInfoList.sort(Comparator.comparingInt(FieldInfo::getOrder));

        return fieldInfoList;
    }

    /**
     * 获取类的所有字段（包括父类字段）
     *
     * @param clazz 类
     * @return 字段列表
     */
    private static List<Field> getAllFields(Class<?> clazz) {
        List<Field> fields = new ArrayList<>();
        Class<?> currentClass = clazz;

        while (currentClass != null && currentClass != Object.class) {
            fields.addAll(Arrays.asList(currentClass.getDeclaredFields()));
            currentClass = currentClass.getSuperclass();
        }

        return fields;
    }


    /**
     * 写入导出时间和金额单位信息行（第二行）
     *
     * @param sheet      工作表
     * @param titleStyle 标题行样式
     */
    private static void writeExportTableTitle(Sheet sheet, int columnCount, String title, CellStyle titleStyle) {

        CellRangeAddress region = new CellRangeAddress(0, 0, 0, columnCount - 1);
        sheet.addMergedRegion(region);
        Row infoRow = sheet.createRow(0);

        // 第二行第一列：导出时间
        Cell timeCell = infoRow.createCell(0);
        timeCell.setCellValue(title);
        timeCell.setCellStyle(titleStyle);


    }

    /**
     * 写入导出时间和金额单位信息行（第二行）
     *
     * @param sheet       工作表
     * @param columnCount 列数
     * @param moneyUnit   金额单位
     * @param infoStyle   信息行样式
     */
    private static void writeExportInfoRow(Sheet sheet, int columnCount, String moneyUnit, CellStyle infoStyle) {
        Row infoRow = sheet.createRow(1);

        // 第二行第一列：导出时间
        Cell timeCell = infoRow.createCell(0);
        String exportTime = "导出时间：" + LocalDateTime.now().format(DEFAULT_DATE_FORMATTER);
        timeCell.setCellValue(exportTime);
        timeCell.setCellStyle(infoStyle);

        // 如果有金额单位，设置在最后一列
        if (moneyUnit != null && !moneyUnit.isEmpty()) {
            Cell unitCell = infoRow.createCell(columnCount - 1);
            String unitInfo = "金额单位：" + moneyUnit;
            unitCell.setCellValue(unitInfo);
            unitCell.setCellStyle(infoStyle);
        }
    }

    /**
     * 写入表头
     *
     * @param sheet         工作表
     * @param fieldInfoList 字段信息列表
     * @param headerStyle   表头样式
     */
    private static void writeHeaders(Sheet sheet, List<FieldInfo> fieldInfoList, CellStyle headerStyle) {
        Row headerRow = sheet.createRow(2);
        headerRow.setHeightInPoints(18); // 设置表头行高

        for (int i = 0; i < fieldInfoList.size(); i++) {
            Cell cell = headerRow.createCell(i);
            cell.setCellValue(fieldInfoList.get(i).getTitle());
            cell.setCellStyle(headerStyle);
        }
    }


    /**
     * 写入数据行
     *
     * @param sheet         工作表
     * @param rowIndex      行索引
     * @param data          数据对象
     * @param fieldInfoList 字段信息列表
     * @param styles        通用样式映射
     * @param fieldStyles   字段自定义样式映射
     * @param sumMap        合计值映射
     * @throws Exception 异常
     */
    private static void writeDataRow(Sheet sheet, int rowIndex, Object data, List<FieldInfo> fieldInfoList,
                                     Map<String, CellStyle> styles, Map<Field, CellStyle> fieldStyles,
                                     Map<Integer, BigDecimal> sumMap) throws Exception {
        Row row = sheet.createRow(rowIndex);

        for (int i = 0; i < fieldInfoList.size(); i++) {
            FieldInfo fieldInfo = fieldInfoList.get(i);
            Field field = fieldInfo.getField();

            // 获取字段值
            Object value = field.get(data);
            if (value == null) {
                continue;
            }

            Cell cell = row.createCell(i);

            // 根据字段类型设置单元格值和样式
            setCellValueAndStyle(cell, value, field, styles, fieldStyles);

            // 如果需要合计，累加值
            if (fieldInfo.isSum() && value instanceof Number) {
                BigDecimal numValue;
                if (value instanceof BigDecimal) {
                    numValue = (BigDecimal) value;
                } else {
                    numValue = new BigDecimal(value.toString());
                }

                // 将值累加到对应的列
                sumMap.merge(i, numValue, BigDecimal::add);
            }
        }
    }

    /**
     * 设置单元格值和样式
     *
     * @param cell        单元格
     * @param value       字段值
     * @param field       字段
     * @param styles      通用样式映射
     * @param fieldStyles 字段自定义样式映射
     */
    private static void setCellValueAndStyle(Cell cell, Object value, Field field,
                                             Map<String, CellStyle> styles, Map<Field, CellStyle> fieldStyles) {
        // 首先检查是否有自定义样式
        CellStyle style = fieldStyles.get(field);

        if (value instanceof String) {
            cell.setCellValue((String) value);
            if (style == null) {
                style = styles.get("default");
            }
        } else if (value instanceof BigDecimal) {
            cell.setCellValue(((BigDecimal) value).doubleValue());
            if (style == null) {
                style = styles.get("number");
            }
        } else if (value instanceof Integer || value instanceof Long) {
            cell.setCellValue(((Number) value).longValue());
            if (style == null) {
                style = styles.get("integer");
            }
        } else if (value instanceof Double || value instanceof Float) {
            cell.setCellValue(((Number) value).doubleValue());
            if (style == null) {
                style = styles.get("number");
            }
        } else if (value instanceof Boolean) {
            if ((Boolean) value) {
                value = "是";
            } else {
                value = "否";
            }
            cell.setCellValue((String) value);
            if (style == null) {
                style = styles.get("default");
            }
        } else if (value instanceof Date) {
            cell.setCellValue((Date) value);
            if (style == null) {
                style = styles.get("date");
            }
        } else if (value instanceof LocalDateTime) {
            Date date = Date.from(((LocalDateTime) value).atZone(ZoneId.systemDefault()).toInstant());
            cell.setCellValue(date);
            if (style == null) {
                style = styles.get("datetime");
            }
        } else if (value instanceof LocalDate) {
            Date date = Date.from(((LocalDate) value).atStartOfDay(ZoneId.systemDefault()).toInstant());
            cell.setCellValue(date);
            if (style == null) {
                style = styles.get("date");
            }
        } else {
            // 其他类型转为字符串
            cell.setCellValue(value.toString());
            if (style == null) {
                style = styles.get("default");
            }
        }

        cell.setCellStyle(style);
    }

    /**
     * 写入合计行
     *
     * @param sheet         工作表
     * @param rowIndex      行索引
     * @param fieldInfoList 字段信息列表
     * @param sumMap        合计值映射
     * @param styles        样式映射
     */
    private static void writeSummaryRow(Sheet sheet, int rowIndex, List<FieldInfo> fieldInfoList,
                                        Map<Integer, BigDecimal> sumMap, Map<String, CellStyle> styles) {
        Row summaryRow = sheet.createRow(rowIndex);

        // 第一列写"合计"
        Cell summaryLabelCell = summaryRow.createCell(0);
        summaryLabelCell.setCellValue("合计");
        summaryLabelCell.setCellStyle(styles.get("summaryLabel"));

        // 写入合计值
        for (Map.Entry<Integer, BigDecimal> entry : sumMap.entrySet()) {
            int colIndex = entry.getKey();
            BigDecimal sum = entry.getValue();

            Cell summaryValueCell = summaryRow.createCell(colIndex);
            summaryValueCell.setCellValue(sum.doubleValue());
            summaryValueCell.setCellStyle(styles.get("summaryNumber"));
        }
    }

    /**
     * 字段信息类
     */
    private static class FieldInfo {
        private final Field field;
        private final String title;
        private final int order;
        private final String dateFormat;
        private final String numberFormat;
        private final boolean sum;
        private final int width;

        public FieldInfo(Field field, String title, int order, String dateFormat, String numberFormat, boolean sum, int width) {
            this.field = field;
            this.title = title;
            this.order = order;
            this.dateFormat = dateFormat;
            this.numberFormat = numberFormat;
            this.sum = sum;
            this.width = width;
        }

        public Field getField() {
            return field;
        }

        public String getTitle() {
            return title;
        }

        public int getOrder() {
            return order;
        }

        public String getDateFormat() {
            return dateFormat;
        }

        public String getNumberFormat() {
            return numberFormat;
        }

        public boolean isSum() {
            return sum;
        }

        public int getWidth() {
            return width;
        }
    }

    /**
     * Sheet数据类
     *
     * @param <T> 数据类型
     */
    public static class SheetData<T> {
        private final List<T> dataList;
        private final Class<T> clazz;

        public SheetData(List<T> dataList, Class<T> clazz) {
            this.dataList = dataList;
            this.clazz = clazz;
        }

        public List<T> getDataList() {
            return dataList;
        }

        public Class<T> getClazz() {
            return clazz;
        }
    }
}

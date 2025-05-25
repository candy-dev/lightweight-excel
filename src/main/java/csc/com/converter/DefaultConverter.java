package csc.com.converter;

import csc.com.annotation.ExcelColumn;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.DateUtil;

import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;

/**
 * 默认转换器
 * @author 刺猬
 */
public class DefaultConverter implements DataConverter<Object>{
    @Override
    public Object convert(Object cellValue, Field field) throws Exception {
        if (cellValue == null ) {
            return null;
        }

        String value = getCellValueAsString(cellValue);
        if (StringUtils.isBlank(value)) {
            return null;
        }

        Class<?> fieldType = field.getType();

//        if (fieldType == String.class) {
//            return value.trim();
//        } else if (fieldType == Integer.class || fieldType == int.class) {
//            return Integer.valueOf(value.trim());
//        } else if (fieldType == Long.class || fieldType == long.class) {
//            return Long.valueOf(value.trim());
//        } else if (fieldType == Double.class || fieldType == double.class) {
//            return Double.valueOf(value.trim());
//        } else if (fieldType == BigDecimal.class) {
//            return new BigDecimal(value.trim());
//        } else if (fieldType == Boolean.class || fieldType == boolean.class) {
//            return Boolean.valueOf(value.trim());
//        } else if (fieldType == LocalDate.class) {
//            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
//            String pattern = annotation != null ? annotation.dateFormat() : "yyyy-MM-dd";
//            return LocalDate.parse(value.trim(), DateTimeFormatter.ofPattern(pattern));
//        } else if (fieldType == LocalDateTime.class) {
//            ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
//            String pattern = annotation != null ? annotation.dateFormat() : "yyyy-MM-dd HH:mm:ss";
//            return LocalDateTime.parse(value.trim(), DateTimeFormatter.ofPattern(pattern));
//        }

        if (fieldType == String.class) {
            return value.trim();
        } else if (fieldType == Integer.class || fieldType == int.class) {
            return parseInteger(value);
        } else if (fieldType == Long.class || fieldType == long.class) {
            return parseLong(value);
        } else if (fieldType == Double.class || fieldType == double.class) {
            return parseDouble(value);
        } else if (fieldType == BigDecimal.class) {
            return parseBigDecimal(value);
        } else if (fieldType == Boolean.class || fieldType == boolean.class) {
            return parseBoolean(value);
        } else if (fieldType == LocalDate.class) {
            return parseLocalDate(cellValue, field);
        } else if (fieldType == LocalDateTime.class) {
            return parseLocalDateTime(cellValue, field);
        }


        return value.trim();
    }

    private String getCellValueAsString(Object cellValue) {
        if (cellValue instanceof String) {
            return (String) cellValue;
        } else if (cellValue instanceof Number) {
            // 如果是整数，不显示小数点
            Number num = (Number) cellValue;
            if (num.doubleValue() == num.intValue()) {
                return String.valueOf(num.intValue());
            }
            return num.toString();
        } else if (cellValue instanceof Date) {
            return DateTimeFormatter.ofPattern("yyyy-MM-dd HH:mm:ss")
                    .format(((Date) cellValue).toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime());
        } else if (cellValue instanceof Boolean) {
            return cellValue.toString();
        }
        return cellValue.toString();
    }
    private Integer parseInteger(String value) {
        try {
            // 处理可能的小数点
            if (value.contains(".")) {
                return Double.valueOf(value.trim()).intValue();
            }
            return Integer.valueOf(value.trim());
        } catch (NumberFormatException e) {
            throw new RuntimeException("无法转换为整数: " + value);
        }
    }

    private Long parseLong(String value) {
        try {
            if (value.contains(".")) {
                return Double.valueOf(value.trim()).longValue();
            }
            return Long.valueOf(value.trim());
        } catch (NumberFormatException e) {
            throw new RuntimeException("无法转换为长整数: " + value);
        }
    }

    private Double parseDouble(String value) {
        try {
            return Double.valueOf(value.trim());
        } catch (NumberFormatException e) {
            throw new RuntimeException("无法转换为小数: " + value);
        }
    }

    private BigDecimal parseBigDecimal(String value) {
        try {
            return new BigDecimal(value.trim());
        } catch (NumberFormatException e) {
            throw new RuntimeException("无法转换为BigDecimal: " + value);
        }
    }

    private Boolean parseBoolean(String value) {
        String trimmed = value.trim().toLowerCase();
        if ("true".equals(trimmed) || "1".equals(trimmed) || "是".equals(trimmed) || "yes".equals(trimmed)) {
            return true;
        } else if ("false".equals(trimmed) || "0".equals(trimmed) || "否".equals(trimmed) || "no".equals(trimmed)) {
            return false;
        }
        throw new RuntimeException("无法转换为布尔值: " + value);
    }

    private LocalDate parseLocalDate(Object cellValue, Field field) {
        if (cellValue instanceof Date) {
            return ((Date) cellValue).toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
        }

        // 处理数值类型的日期（Excel日期序列号）
        if (cellValue instanceof Number) {
            try {
                Date date = DateUtil.getJavaDate(((Number) cellValue).doubleValue());
                return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDate();
            } catch (Exception e) {
                // 如果不是日期序列号，按字符串处理
            }
        }

        String value = getCellValueAsString(cellValue);
        if (StringUtils.isBlank(value)) {
            return null;
        }

        ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
        String pattern = annotation != null ? annotation.dateFormat() : "yyyy-MM-dd";

        // 尝试多种日期格式
        String[] patterns = {
                pattern,
                "yyyy-MM-dd",
                "yyyy/MM/dd",
                "yyyy-M-d",
                "yyyy/M/d",
                "dd/MM/yyyy",
                "dd-MM-yyyy",
                "MM/dd/yyyy",
                "MM-dd-yyyy",
                "yyyyMMdd"
        };

        for (String p : patterns) {
            try {
                return LocalDate.parse(value.trim(), DateTimeFormatter.ofPattern(p));
            } catch (Exception ignored) {
                // 继续尝试下一个格式
            }
        }

        throw new RuntimeException("日期格式错误，支持格式: " + String.join(", ", patterns) + ", 实际值: " + value);
    }

    private LocalDateTime parseLocalDateTime(Object cellValue, Field field) {
        if (cellValue instanceof Date) {
            return ((Date) cellValue).toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
        }

        // 处理数值类型的日期（Excel日期序列号）
        if (cellValue instanceof Number) {
            try {
                Date date = DateUtil.getJavaDate(((Number) cellValue).doubleValue());
                return date.toInstant().atZone(ZoneId.systemDefault()).toLocalDateTime();
            } catch (Exception e) {
                // 如果不是日期序列号，按字符串处理
            }
        }

        String value = getCellValueAsString(cellValue);
        if (StringUtils.isBlank(value)) {
            return null;
        }

        ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
        String pattern = annotation != null ? annotation.dateFormat() : "yyyy-MM-dd HH:mm:ss";

        // 尝试多种日期时间格式
        String[] patterns = {
                pattern,
                "yyyy-MM-dd HH:mm:ss",
                "yyyy-MM-dd HH:mm",
                "yyyy/MM/dd HH:mm:ss",
                "yyyy/MM/dd HH:mm",
                "yyyy-M-d H:m:s",
                "yyyy/M/d H:m:s",
                "dd/MM/yyyy HH:mm:ss",
                "dd-MM-yyyy HH:mm:ss",
                "MM/dd/yyyy HH:mm:ss",
                "MM-dd-yyyy HH:mm:ss",
                "yyyyMMdd HHmmss",
                "yyyy-MM-dd"  // 如果只有日期部分，时间默认为00:00:00
        };

        for (String p : patterns) {
            try {
                if (p.equals("yyyy-MM-dd")) {
                    // 只有日期的情况，转换为LocalDateTime
                    LocalDate date = LocalDate.parse(value.trim(), DateTimeFormatter.ofPattern(p));
                    return date.atStartOfDay();
                } else {
                    return LocalDateTime.parse(value.trim(), DateTimeFormatter.ofPattern(p));
                }
            } catch (Exception ignored) {
                // 继续尝试下一个格式
            }
        }

        throw new RuntimeException("日期时间格式错误，支持格式: " + String.join(", ", patterns) + ", 实际值: " + value);
    }
}

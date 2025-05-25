package csc.com.converter;

import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.DateUtil;

import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;

/**
 * 自定义日期转换器示例
 * @author 刺猬
 */
public class DateConverter implements DataConverter<LocalDate> {
    @Override
    public LocalDate convert(Object cellValue, Field field) throws Exception {
        if (cellValue == null) {
            return null;
        }
        
        // 处理Excel原生日期
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
        
        // 处理字符串日期
        String value = cellValue.toString().trim();
        if (StringUtils.isBlank(value)) {
            return null;
        }
        
        // 尝试多种日期格式
        String[] patterns = {
            "yyyy-MM-dd",
            "yyyy/MM/dd",
            "dd/MM/yyyy",
            "MM/dd/yyyy",
            "yyyyMMdd",
            "yyyy年MM月dd日",
            "yyyy-M-d",
            "yyyy/M/d"
        };
        
        for (String pattern : patterns) {
            try {
                return LocalDate.parse(value, DateTimeFormatter.ofPattern(pattern));
            } catch (Exception ignored) {
                // 继续尝试下一个格式
            }
        }
        
        throw new RuntimeException("无法解析日期格式: " + value + ", 支持的格式: " + String.join(", ", patterns));
    }
}
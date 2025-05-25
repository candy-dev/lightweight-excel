package csc.com.annotation;

import csc.com.converter.DataConverter;
import csc.com.converter.DefaultConverter;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * 导入注解
 * @author 刺猬
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {
    /**
     * Excel列名
     */
    String value();

    /**
     * 列索引（从0开始）
     */
    int index() default -1;

    /**
     * 是否必填
     */
    boolean required() default false;

    /**
     * 日期格式
     */
    String dateFormat() default "yyyy-MM-dd";

    /**
     * 默认值
     */
    String defaultValue() default "";

    /**
     * 是否为数值列（用于合计行判断）
     */
    boolean isNumeric() default false;
    /**
     * 数值格式
     */
    String numberFormat() default "";
    /**
     * 数据转换器
     */
    Class<? extends DataConverter> converter() default DefaultConverter.class;
}

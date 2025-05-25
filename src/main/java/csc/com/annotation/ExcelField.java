package csc.com.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel导出字段注解
 * 用于标记需要导出到Excel的字段
 * @author 刺猬
 */
@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelField {

    /**
     * 字段标题(Excel列标题)
     */
    String title();

    /**
     * 字段排序序号(决定列的顺序，从小到大排序)
     */
    int order() default 0;

    /**
     * 日期格式，如: yyyy-MM-dd
     */
    String dateFormat() default "";

    /**
     * 数值格式，如: #,##0.00
     */
    String numberFormat() default "";

    /**
     * 是否需要合计，默认false
     */
    boolean sum() default false;

    /**
     * 宽度
     */
    int width() default 20;

    /**
     * 是否忽略该字段
     */
    boolean ignore() default false;
}

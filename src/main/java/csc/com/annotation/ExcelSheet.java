package csc.com.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * Excel导出Sheet注解
 * 用于标记需要导出为Excel工作表的类
 * @author 刺猬
 */
@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelSheet {
    /**
     * 工作表名称
     */
    String name() default "";

    /**
     * 金额单位
     */
    String moneyUnit() default "";
    /**
     * Sheet索引
     */
    int index() default -1;

    /**
     * 标题行索引（从0开始，默认第3行）
     */
    int headerRow() default 2;

    /**
     * 数据开始行索引（默认第4行）
     */
    int dataStartRow() default 3;

    /**
     * 是否跳过合计行
     */
    boolean skipSummaryRow() default true;
}

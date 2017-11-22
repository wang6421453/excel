package cn.ijiami.framework.excel.parser;

import java.lang.annotation.*;

/**
 * excel行注解
 * 
 * @author wl
 * @date 2017/11/8
 */

@Target(ElementType.TYPE)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface Excel {

    /**标题行位置，默认第一行（1）*/
	int titleRowIndex() default 1;

	/**内容行索引，默认第二行（2）*/
	int dataRowIndex() default 2;

	/**sheet页索引，默认第一页（1）*/
	int sheetIndex() default 1;

	/**sheet页名称，默认（Sheet1）*/
	String sheetName() default "Sheet1";
}

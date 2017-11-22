package cn.ijiami.framework.excel.parser;

import java.lang.annotation.*;

/**
 * excel列注解
 * 
 * @author wl
 * @date 2017/11/8
 */

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Documented
public @interface ExcelColumn {

	/**列索引*/
	int value() default 1;

	/**列名称*/
	String columnName() default "";
}

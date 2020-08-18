package com.jeesite.common.utils.excel.fieldtype;

import java.lang.annotation.*;

/**
 * @description:通过此注解，可以获取类属性在Excel的位置
 * @author: Mr.Luke
 * @create: 2019-07-16 17:02
 * @Version V1.0
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.FIELD, ElementType.TYPE })
public @interface ExcelCell {

    int value() default 0;
}

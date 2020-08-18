package com.jeesite.common.utils.excel.fieldtype;

import java.lang.annotation.*;

/**
 * @description: 属性在Excel的位置
 * @author: Mr.Luke
 * @create: 2019-07-22 11:27
 * @Version V1.0
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.FIELD, ElementType.TYPE })
public @interface  ExcelCoordinate {

    int row() default 0;

    int col() default 0;
}

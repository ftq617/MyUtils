package com.jeesite.common.utils.excel.fieldtype;

import java.lang.annotation.*;

/**
 * @description:通过此注解，可以获取List属性 在第几行开始填充数值
 * @author: Mr.Luke
 * @create: 2019-07-16 17:02
 * @Version V1.0
 */
@Documented
@Retention(RetentionPolicy.RUNTIME)
@Target({ ElementType.FIELD, ElementType.TYPE })
public @interface ExcelListStart {

    int value() default 0;
}

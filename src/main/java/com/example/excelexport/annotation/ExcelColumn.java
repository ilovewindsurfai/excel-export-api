package com.example.excelexport.annotation;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelColumn {
    String name() default "";
    String messageKey() default "";
    int order() default 0;
    String dateFormat() default "yyyy-MM-dd";
    String amountFormat() default "#,##0.00";
    boolean isAmount() default false;
}

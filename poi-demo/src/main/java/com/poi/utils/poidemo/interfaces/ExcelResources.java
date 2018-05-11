package com.poi.utils.poidemo.interfaces;

import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;

/**
 * 声明注解  在对应的类上进行标注 属性为该列的名称与显示顺序
 */
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelResources {

    /**
     * 属性的标题名称
     * @return
     */
    String title();
    /**
     * 在excel的顺序
     * @return
     */
    int order() default 9999;
}

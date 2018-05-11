package com.poi.utils.poidemo.utils;

/**
 * excel 的表头封装类
 * 可以决定表头的名称与顺序显示
 */
public class ExcelHeader implements Comparable<ExcelHeader>{
    /**
     * excel的标题名称
     */
    private String title;
    /**
     * 每一个标题的顺序
     */
    private int order;
    /**
     * 对应方法名称
     */
    private String methodName;

    public String getTitle() {
        return title;
    }
    public void setTitle(String title) {
        this.title = title;
    }
    public int getOrder() {
        return order;
    }
    public void setOrder(int order) {
        this.order = order;
    }
    public String getMethodName() {
        return methodName;
    }
    public void setMethodName(String methodName) {
        this.methodName = methodName;
    }

    public int compareTo(ExcelHeader o) {
        return order>o.order?1:(order<o.order?-1:0);
    }
    public ExcelHeader(String title, int order, String methodName) {
        super();
        this.title = title;
        this.order = order;
        this.methodName = methodName;
    }
}

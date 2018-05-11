package com.poi.utils.poidemo.utils;

import com.monitorjbl.xlsx.StreamingReader;
import com.poi.utils.poidemo.interfaces.ExcelResources;
import org.apache.commons.beanutils.BeanUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.OLE2NotOfficeXmlFileException;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.net.URLEncoder;
import java.util.*;

/**
 * poi  excel 工具类封装
 */
public class ExcelUtil {

    private volatile static ExcelUtil eu = null;
    //双重锁单例模式
    public static ExcelUtil getInstance() {
        if(eu == null){
            synchronized (ExcelUtil.class){
                if(eu == null){
                    eu = new ExcelUtil();
                }
            }
        }
        return eu;
    }

    /**
     * 转换为excel
     * @param objs 数据
     * @param clz 表头对应的类
     * @return
     */
    private  Workbook handleExcel(List objs, Class clz,Boolean isXlsx) {
        //创建工作簿
        Workbook wb = null;
        //poi对于.xls与.xlsx的处理方式有所区分
        if(isXlsx){
            wb = new XSSFWorkbook();
        }else{
            wb = new HSSFWorkbook();
        }

        try {
            //创建一个工作表
            Sheet sheet = wb.createSheet();
            sheet.autoSizeColumn(0);
            //创建第零行——作为表头使用
            Row r = sheet.createRow(0);
            //创建表头
            List<ExcelHeader> headers = getHeaderList(clz);

            //利用集合对list进行排序
            Collections.sort(headers);
            //写标题
            for(int i=0;i<headers.size();i++) {
                sheet.autoSizeColumn(i);
                // 设置字体
                Font font = wb.createFont();
                font.setFontHeightInPoints((short) 13); //字体高度
                font.setColor(Font.COLOR_NORMAL); //字体颜色
                font.setFontName("黑体"); //字体

                //设置单元格的样式
                CellStyle style = wb.createCellStyle();
                style.setFont(font);
                style.setWrapText(true);
                style.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
                style.setAlignment(HorizontalAlignment.CENTER);
                style.setVerticalAlignment(VerticalAlignment.CENTER);

                //根据表头的数量填充第零行
                r.createCell(i).setCellValue(headers.get(i).getTitle());
                r.getCell(i).setCellStyle(style);
            }
            //写数据
            Object obj = null;
            for(int i=0;i<objs.size();i++) {
                //从第一行起
                r = sheet.createRow(i+1);
                obj = objs.get(i);
                for(int j=0;j<headers.size();j++) {
                    //给方法为通过名称获取obj对象中的值，然后依照顺序（上面已经对list进行过排序）依次填充数据即可
                    String value = BeanUtils.getProperty(obj, getMethodName(headers.get(j)));
                    r.createCell(j).setCellValue(value);
                    //设置行宽
                    int herderWidth = headers.get(j).getTitle().getBytes().length * 2 * 150;
                    int dataWidth = value.getBytes().length * 2 * 150;
                    sheet.setColumnWidth(j, dataWidth>herderWidth?dataWidth:herderWidth);
                }
            }
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        } catch (NoSuchMethodException e) {
            e.printStackTrace();
        }
        return wb;
    }

    /**
     * 利用反射获取类当中的方法，从而获取应该创建的表头的信息
     * @param clz
     * @return
     */
    private  List<ExcelHeader> getHeaderList(Class clz) {
        List<ExcelHeader> headers = new ArrayList<ExcelHeader>();
        Method[] ms = clz.getDeclaredMethods();
        for(Method m:ms) {
            String mn = m.getName();
            if(mn.startsWith("get")) {
                //判断该类的get方法上是否有自己生声明的注解，从而获取表头显示的中文名字以及顺序
                if(m.isAnnotationPresent(ExcelResources.class)) {
                    ExcelResources er = m.getAnnotation(ExcelResources.class);
                    headers.add(new ExcelHeader(er.title(),er.order(),mn));
                }
            }
        }
        return headers;
    }

    /**
     * 根据标题获取相应的方法名称
     * @param eh
     * @return
     */
    private  String getMethodName(ExcelHeader eh) {
        //去掉get和set方法
        String mn = eh.getMethodName().substring(3);
        //同意转换为小写
        mn = mn.substring(0,1).toLowerCase()+mn.substring(1);
        return mn;
    }

    /*-----------------------------------------------------这是一条分割线--------------------------------------------*/
    /*-------------------------------------以下的方法为实际调用的方法------------------------------------------------*/

    /**
     * 基于流导出对象到excel
     * @param response 请求返回
     * @param objs 对象列表
     * @param clz 需要转换成表头的类
     * @param fileName 导出后的文件名（不包含包含后缀 .xls 或者.xlsx）
     * @param  isXlsx   是否后缀为.xlsx文件     false 为 .xls
     */
    public  void exportExcelByStream(HttpServletResponse response, List objs, Class clz, String fileName,Boolean isXlsx) {
        if(isXlsx){
            fileName = fileName+".xlsx";
        }else{
            fileName = fileName+".xls";
        }
        try {
            // 告诉浏览器用什么软件可以打开此文件
            response.setHeader("content-Type", "application/vnd.ms-excel");
            // 下载文件的默认名称
            response.setHeader("Content-Disposition", "attachment;filename="+URLEncoder.encode(fileName, "utf-8"));

            Workbook wb = handleExcel(objs, clz,isXlsx);
            wb.write(response.getOutputStream());
        } catch (IOException e) {
            e.printStackTrace();
        }
    }



    public void  readExcel(String filePath, InputStream inp) throws IOException {
        Workbook wb;
        if (filePath.endsWith(".xls")) {
            wb = new HSSFWorkbook(inp);
        } else if (filePath.endsWith(".xlsx")) {
            wb = new XSSFWorkbook(inp);
        } else {
            wb = StreamingReader.builder()
                    .rowCacheSize(1000)    // number of rows to keep in memory (defaults to 10)
                    .bufferSize(4096)     // buffer size to use when reading InputStream to file (defaults to 1024)
                    .open(inp);            // InputStream or File for XLSX file (required)
        }
        Sheet sheet = wb.getSheetAt(0);
        //遍历所有的行
        for (Row row : sheet) {
            System.out.println("开始遍历第" + row.getRowNum() + "行数据：");
            //遍历所有的列
            for (Cell cell : row) {
                System.out.print(cell.getStringCellValue() + " ");
            }
            System.out.println(" ");
        }
    }

}
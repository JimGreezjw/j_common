import fr.opensagres.poi.xwpf.converter.pdf.PdfConverter;
import fr.opensagres.poi.xwpf.converter.pdf.PdfOptions;
import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFDocument;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;

public class POIutils {
    public static void poiUtils(Object o, List list) throws Exception {
        //获取对象的类对象
        Class<?> aClass = o.getClass();
        //创建一个List，用来存放对象属性的get方法名
        List<String> methodNameList = new ArrayList<>();
        //获取对象的所有属性
        Field[] declaredFields = aClass.getDeclaredFields();
        //遍历获取所有属性的get方法名并存在一个数组中
        for (Field declaredField : declaredFields) {
            String fileName = declaredField.getName();
            String getMethodName = "get" + fileName.substring(0, 1).toUpperCase() + fileName.substring(1);
            System.out.println(getMethodName);
            methodNameList.add(getMethodName);
        }


        HSSFWorkbook sheets = new HSSFWorkbook();
        HSSFSheet sheet = sheets.createSheet(aClass.getName());
        HSSFRow row = sheet.createRow(0);
        HSSFCell cell = null;
        for (int i = 0; i < declaredFields.length; i++) {
            cell = row.createCell(i);
            cell.setCellValue(declaredFields[i].toString());

        }
        for (int i = 0; i < list.size(); i++) {
            row = sheet.createRow(i);
            for (int j = 0; j < declaredFields.length; j++) {

                cell = row.createCell(j);
                Method declaredMethod = aClass.getDeclaredMethod(methodNameList.get(j), null);
                Object invoke = declaredMethod.invoke(list.get(i), null);
                System.out.print(invoke);
                cell.setCellValue(invoke.toString());
            }
            System.out.println();
        }
        String xlsName = aClass.getName();
        //这里涉及到“.”在正则表达式中的特殊性，将在下一篇文章中讨论
        String replace = xlsName.replace(".", "/");
        String[] split = replace.split("/");
        System.out.println(split.length);
        String s = split[split.length - 1];
        sheets.write(new FileOutputStream("E:\\" + s + ".xls"));
    }

}

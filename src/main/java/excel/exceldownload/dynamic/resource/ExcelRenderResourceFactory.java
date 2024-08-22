package excel.exceldownload.dynamic.resource;

import excel.exceldownload.dynamic.annotation.ExcelBody;
import excel.exceldownload.dynamic.annotation.ExcelColumnStyle;
import excel.exceldownload.dynamic.annotation.ExcelHeader;
import excel.exceldownload.dynamic.style.ExcelCellStyle;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.LinkedHashMap;
import java.util.Map;

public class ExcelRenderResourceFactory {

    public static ExcelRenderResource prepareRenderResource(Class<?> clazz, Workbook workbook) throws InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
        Map<String, Integer> columnSizesMap = new LinkedHashMap<>();
        PreCalculatedCellStyleMap styleMap = new PreCalculatedCellStyleMap();

        Field[] fields = clazz.getDeclaredFields();
        for (Field field : fields) {
            if (field.isAnnotationPresent(ExcelHeader.class)) {
                ExcelHeader annotation = field.getAnnotation(ExcelHeader.class);
                columnSizesMap.put(field.getName(), annotation.columnSize());
                styleMap.put(field.getName(), getCellStyle(annotation.style()), workbook);
            }

            if (field.isAnnotationPresent(ExcelBody.class)) {
                ExcelBody annotation = field.getAnnotation(ExcelBody.class);
                styleMap.put(field.getName(), getCellStyle(annotation.style()), workbook);
            }
        }

        return new ExcelRenderResource(columnSizesMap, styleMap);
    }


    private static ExcelCellStyle getCellStyle(ExcelColumnStyle excelColumnStyle) throws NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException {
        Class<? extends ExcelCellStyle> excelCellStyleClass = excelColumnStyle.excelCellStyleClass();

        return excelCellStyleClass.getDeclaredConstructor().newInstance();
    }

}

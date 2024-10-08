package excel.exceldownload.main.excel.resource;

import excel.exceldownload.main.annotation.DefaultBodyStyle;
import excel.exceldownload.main.annotation.DefaultHeaderStyle;
import excel.exceldownload.main.annotation.ExcelColumn;
import excel.exceldownload.main.annotation.ExcelColumnStyle;
import excel.exceldownload.main.excel.resource.collection.PreCalculatedCellStyleMap;
import excel.exceldownload.main.excel.style.ExcelCellStyle;
import excel.exceldownload.main.excel.style.NoExcelCellStyle;
import excel.exceldownload.main.exception.InvalidExcelCellStyleException;
import excel.exceldownload.main.exception.NoExcelColumnAnnotationsException;
import org.apache.poi.ss.usermodel.Workbook;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import static excel.exceldownload.main.utils.ReflectionUtils.getAllFields;
import static excel.exceldownload.main.utils.ReflectionUtils.getAnnotation;

public class ExcelRenderResourceFactory {

    public static ExcelRenderResource prepareRenderResource(Class<?> type, Workbook workbook) {
        PreCalculatedCellStyleMap styleMap = new PreCalculatedCellStyleMap();
        Map<String, String> headerNamesMap = new LinkedHashMap<>();
        Map<String, Integer> columnSizesMap = new LinkedHashMap<>();
        List<String> fieldNames = new ArrayList<>();

        ExcelColumnStyle classDefinedHeaderStyle = getHeaderExcelColumnStyle(type);
        ExcelColumnStyle classDefinedBodyStyle = getBodyExcelColumnStyle(type);

        for (Field field : getAllFields(type)) {
            if (field.isAnnotationPresent(ExcelColumn.class)) {
                ExcelColumn annotation = field.getAnnotation(ExcelColumn.class);
                styleMap.put(
                        ExcelCellKey.of(field.getName(), ExcelRenderLocation.HEADER),
                        getCellStyle(decideAppliedStyleAnnotation(classDefinedHeaderStyle, annotation.headerStyle())), workbook);
                styleMap.put(
                        ExcelCellKey.of(field.getName(), ExcelRenderLocation.BODY),
                        getCellStyle(decideAppliedStyleAnnotation(classDefinedBodyStyle, annotation.bodyStyle())), workbook);
                fieldNames.add(field.getName());
                headerNamesMap.put(field.getName(), annotation.headerName());
                columnSizesMap.put(field.getName(), annotation.columnSize());
            }
        }

        if (styleMap.isEmpty()) {
            throw new NoExcelColumnAnnotationsException(String.format("Class %s has not @ExcelColumn at all", type));
        }
        return new ExcelRenderResource(styleMap, headerNamesMap, columnSizesMap, fieldNames);
    }

    private static ExcelColumnStyle getHeaderExcelColumnStyle(Class<?> clazz) {
        Annotation annotation = getAnnotation(clazz, DefaultHeaderStyle.class);
        if (annotation == null) {
            return null;
        }
        return ((DefaultHeaderStyle) annotation).style();
    }

    private static ExcelColumnStyle getBodyExcelColumnStyle(Class<?> clazz) {
        Annotation annotation = getAnnotation(clazz, DefaultBodyStyle.class);
        if (annotation == null) {
            return null;
        }
        return ((DefaultBodyStyle) annotation).style();
    }

    private static ExcelColumnStyle decideAppliedStyleAnnotation(ExcelColumnStyle classAnnotation,
                                                                 ExcelColumnStyle fieldAnnotation) {
        if (fieldAnnotation.excelCellStyleClass().equals(NoExcelCellStyle.class) && classAnnotation != null) {
            return classAnnotation;
        }
        return fieldAnnotation;
    }

    private static ExcelCellStyle getCellStyle(ExcelColumnStyle excelColumnStyle) {
        Class<? extends ExcelCellStyle> excelCellStyleClass = excelColumnStyle.excelCellStyleClass();

        try {
            return excelCellStyleClass.getDeclaredConstructor().newInstance();
        } catch (InvocationTargetException | InstantiationException | IllegalAccessException | NoSuchMethodException e) {
            throw new InvalidExcelCellStyleException(e.getMessage(), e);
        }
    }
}

package excel.exceldownload.dynamic.annotation;

import excel.exceldownload.dynamic.style.NoExcelCellStyle;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
public @interface ExcelBody {

    ExcelColumnStyle style() default @ExcelColumnStyle(excelCellStyleClass = NoExcelCellStyle.class);

}

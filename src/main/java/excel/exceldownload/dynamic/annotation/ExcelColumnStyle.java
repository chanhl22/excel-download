package excel.exceldownload.dynamic.annotation;

import excel.exceldownload.dynamic.style.ExcelCellStyle;

public @interface ExcelColumnStyle {

	Class<? extends ExcelCellStyle> excelCellStyleClass();

}

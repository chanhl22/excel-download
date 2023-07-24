package excel.exceldownload.annotation;

import excel.exceldownload.excel.style.ExcelCellStyle;

public @interface ExcelColumnStyle {

	Class<? extends ExcelCellStyle> excelCellStyleClass();
}

package excel.exceldownload.main.annotation;

import excel.exceldownload.main.excel.style.ExcelCellStyle;

public @interface ExcelColumnStyle {

	Class<? extends ExcelCellStyle> excelCellStyleClass();
}

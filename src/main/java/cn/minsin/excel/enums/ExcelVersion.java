package cn.minsin.excel.enums;

import lombok.Getter;
import lombok.RequiredArgsConstructor;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * @author: minton.zhang
 * @since: 2020/3/31 22:57
 */
@Getter
@RequiredArgsConstructor
public enum ExcelVersion {
	/**
	 * 2003版本
	 */
	VERSION_2003(".xls", HSSFWorkbook.class),
	/**
	 * 2007版本
	 */
	VERSION_2007(".xlsx", XSSFWorkbook.class),
	/**
	 * 超过5000条数据
	 */
	VERSION_MORE_THAN_5000(".xlsx", SXSSFWorkbook.class);

	private final String suffix;

	private final Class<?> clazz;


	/**
	 * 判断当前work版本
	 *
	 * @param workbook
	 * @return
	 */
	public static ExcelVersion checkVersion(Workbook workbook) {
		return workbook instanceof HSSFWorkbook ? ExcelVersion.VERSION_2003 :
				workbook instanceof XSSFWorkbook ? ExcelVersion.VERSION_2007 :
						ExcelVersion.VERSION_MORE_THAN_5000;
	}

	public Workbook createInstance() {
		try {
			return (Workbook) this.getClazz().newInstance();
		} catch (Exception e) {
			return null;
		}
	}
}

package cn.minsin.excel.tools;

import cn.minsin.core.tools.StringUtil;
import lombok.Cleanup;
import lombok.NonNull;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import javax.servlet.ServletOutputStream;
import javax.servlet.http.HttpServletResponse;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Iterator;
import java.util.function.Consumer;
import java.util.function.Function;

/**
 * @author: minton.zhang
 * @since: 2020/3/31 22:58
 */
@Slf4j
public class ExcelUtil {


	/**
	 * 获取Excel模板文件
	 *
	 * @param excelPath
	 * @param isInDisk  当为false时 则获取resource下的文件 填写时开头不能以/开头
	 * @return
	 */
	public static InputStream getExcelTemplate(String excelPath, boolean isInDisk) throws FileNotFoundException {
		if (isInDisk) {
			return new FileInputStream(excelPath);
		} else {
			return ExcelUtil.class.getClassLoader().getResourceAsStream(excelPath);
		}
	}

	/**
	 * 获取CellValue
	 *
	 * @param cell
	 */
	public static String getCellRealValue(Cell cell) {
		if (cell == null) {
			return null;
		}
		String cellValue;
		// 以下是判断数据的类型
		switch (cell.getCellType()) {
			// 数字
			case NUMERIC:
				if (DateUtil.isCellDateFormatted(cell)) {
					SimpleDateFormat sdf = new SimpleDateFormat("yyyy-MM-dd");
					cellValue = sdf.format(DateUtil.getJavaDate(cell.getNumericCellValue()));
				} else {
					// 纯数字
					DataFormatter dataFormatter = new DataFormatter();
					cellValue = dataFormatter.formatCellValue(cell);
				}
				break;
			// 字符串
			case STRING:
				cellValue = cell.getStringCellValue();
				break;
			// Boolean
			case BOOLEAN:
				cellValue = cell.getBooleanCellValue() + "";
				break;
			// 公式
			case FORMULA:
				cellValue = cell.getCellFormula() + "";
				break;
			// 空值
			case BLANK:
				cellValue = "";
				break;
			// 故障
			case ERROR:
				cellValue = "非法字符";
				break;
			default:
				cellValue = "未知类型";
				break;
		}
		return StringUtil.filterSpace(cellValue);
	}

	/**
	 * 解析文件输入流
	 *
	 * @param inputStream
	 * @return
	 * @throws IOException
	 */
	public static Workbook parseInputStreamToWorkBook(InputStream inputStream) throws IOException {
		return WorkbookFactory.create(inputStream);
	}

	/**
	 * 解析excel文件
	 *
	 * @param file
	 * @return
	 * @throws IOException
	 */
	public static Workbook parseFileToWorkBook(File file) throws IOException {
		return WorkbookFactory.create(file);
	}

	/**
	 * 行复制功能
	 *
	 * @param fromRow       从哪行开始
	 * @param toRow         目标行
	 * @param copyValueFlag true则连同cell的内容一起复制
	 */
	public static void copyRow(Row fromRow, Row toRow, boolean copyValueFlag) {
		toRow.setHeight(fromRow.getHeight());
		Workbook workbook = fromRow.getSheet().getWorkbook();
		for (Iterator<Cell> cellIt = fromRow.cellIterator(); cellIt.hasNext(); ) {
			Cell tmpCell = cellIt.next();
			Cell newCell = toRow.createCell(tmpCell.getColumnIndex());
			copyCell(workbook, tmpCell, newCell, copyValueFlag);
		}
	}

	/**
	 * 复制单元格 不会进行单元格合并相关
	 *
	 * @param srcCell
	 * @param distCell
	 * @param copyValueFlag true则连同cell的内容一起复制
	 */
	public static void copyCell(Workbook wb, Cell srcCell, Cell distCell, boolean copyValueFlag) {
		CellStyle newStyle = wb.createCellStyle();
		CellStyle srcStyle = srcCell.getCellStyle();

		newStyle.cloneStyleFrom(srcStyle);
		newStyle.setFont(wb.getFontAt(srcStyle.getFontIndexAsInt()));

		// 样式
		distCell.setCellStyle(newStyle);

		// 内容
		if (srcCell.getCellComment() != null) {
			distCell.setCellComment(srcCell.getCellComment());
		}
// 不同数据类型处理
		CellType cellType = srcCell.getCellType();

		distCell.setCellType(cellType);
		if (copyValueFlag) {
			if (cellType == CellType.NUMERIC) {
				if (HSSFDateUtil.isCellDateFormatted(srcCell)) {
					distCell.setCellValue(srcCell.getDateCellValue());
				} else {
					distCell.setCellValue(srcCell.getNumericCellValue());
				}
			} else if (cellType == CellType.STRING) {
				distCell.setCellValue(srcCell.getRichStringCellValue());
			} else if (cellType == CellType.BOOLEAN) {
				distCell.setCellValue(srcCell.getBooleanCellValue());
			} else if (cellType == CellType.ERROR) {
				distCell.setCellErrorValue(srcCell.getErrorCellValue());
			} else if (cellType == CellType.FORMULA) {
				distCell.setCellFormula(srcCell.getCellFormula());
			} else { // nothing29
				distCell.setCellValue(srcCell.getStringCellValue());
			}
		}
	}

	/**
	 * 设置Cell value
	 *
	 * @param row
	 * @param cellIndex
	 * @param cellType
	 * @param value
	 * @param function
	 */
	public static void setCellValue(Row row, int cellIndex, CellType cellType, Object value, Consumer<CellStyle> function) {
		Cell cell = row.getCell(cellIndex);
		if (cell == null) {
			cell = row.createCell(cellIndex);
		}
		cell.setCellType(cellType);
		if (function != null) {
			CellStyle cellStyle = row.getSheet().getWorkbook().createCellStyle();
			//水平居中
			cellStyle.setAlignment(HorizontalAlignment.CENTER);
			//垂直居中
			cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
			//自动换行
			cellStyle.setWrapText(true);
			function.accept(cellStyle);
			cell.setCellStyle(cellStyle);
		}

		if (value == null) {
			cell.setCellValue("");
		} else if (value instanceof String) {
			cell.setCellValue((String) value);
		} else if (value instanceof Integer) {
			cell.setCellValue((Integer) value);
		} else if (value instanceof Boolean) {
			cell.setCellValue((Boolean) value);
		} else if (value instanceof Calendar) {
			cell.setCellValue((Calendar) value);
		} else if (value instanceof Date) {
			cell.setCellValue((Date) value);
		} else {
			cell.setCellValue(value.toString());
		}
	}


	/**
	 * 将workbook转换成为InputStream 并且提供转换转换函数
	 *
	 * @param workbook
	 * @param convert
	 * @param <T>
	 * @return
	 * @throws Exception
	 */
	public static <T> T workbookToInputStream(@NonNull final Workbook workbook, @NonNull final Function<InputStream, T> convert) throws Exception {
		try {
			@Cleanup
			ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();
			workbook.write(byteOutputStream);
			@Cleanup
			ByteArrayInputStream inputStream = new ByteArrayInputStream(byteOutputStream.toByteArray());
			return convert.apply(inputStream);
		} finally {
			if (workbook != null) workbook.close();
		}
	}

	/**
	 * 将workbook转换成为InputStream 并且提供转换函数 常用的例子是上传到oss
	 *
	 * @return
	 * @throws Exception
	 */
	public static void workbookToInputStream(@NonNull final Workbook workbook, @NonNull final Consumer<InputStream> convert) throws Exception {
		try {
			@Cleanup
			ByteArrayOutputStream byteOutputStream = new ByteArrayOutputStream();
			workbook.write(byteOutputStream);
			@Cleanup
			ByteArrayInputStream inputStream = new ByteArrayInputStream(byteOutputStream.toByteArray());
			convert.accept(inputStream);
		} finally {
			if (workbook != null) {
				workbook.close();
			}
		}
	}

	/**
	 * 将workbook转换成为InputStream 并且提供转换函数 常用的例子是上传到oss
	 *
	 * @return
	 * @throws Exception
	 */
	public static void workbookToServletResponse(@NonNull final Workbook workbook,
	                                             @NonNull final HttpServletResponse httpServletResponse,
	                                             @NonNull final String fileName,
	                                             final Function<String, String> nameFunction) throws Exception {
		String newName = nameFunction == null ? fileName : new String(fileName.getBytes(StandardCharsets.UTF_8), StandardCharsets.ISO_8859_1);
		workbookToServletResponse(workbook, httpServletResponse, newName);
	}

	/**
	 * 将workbook转换成为InputStream 并且提供转换函数 常用的例子是上传到oss
	 *
	 * @return
	 * @throws Exception
	 */
	public static void workbookToServletResponse(@NonNull final Workbook workbook,
	                                             @NonNull final HttpServletResponse httpServletResponse,
	                                             @NonNull final String fileName) throws Exception {
		try {
			httpServletResponse.setCharacterEncoding(StandardCharsets.UTF_8.name());
			httpServletResponse.setContentType("application/x-msdownload; charset=utf-8");
			httpServletResponse.setHeader("content-disposition", "attachment;filename=" + fileName);
			ServletOutputStream outputStream = httpServletResponse.getOutputStream();
			workbook.write(outputStream);
		} finally {
			if (workbook != null) workbook.close();
		}
	}
}

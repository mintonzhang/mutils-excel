package cn.minsin.excel.function;

import cn.minsin.core.exception.MutilsException;
import cn.minsin.excel.model.ExcelCellModel;
import cn.minsin.excel.model.ExcelParseResultModel;
import cn.minsin.excel.model.ExcelRowModel;
import cn.minsin.excel.model.ExcelSheetModel;
import cn.minsin.excel.model.SheetParseRule;
import cn.minsin.excel.tools.ExcelUtil;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

import java.io.File;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author: minton.zhang
 * @since: 2020/3/31 23:00
 */
public class ExcelParseFunctions {


	protected Workbook workbook;

	protected List<SheetParseRule> sheetRules = new ArrayList<>(2);

	public ExcelParseFunctions() {
	}

	/**
	 * 静态初始化
	 *
	 * @param file
	 * @return
	 * @throws IOException
	 */
	public static ExcelParseFunctions initByFile(File file) throws IOException {
		return new ExcelParseFunctions().init(file);
	}

	/**
	 * 静态初始化
	 *
	 * @param inputStream
	 * @return
	 * @throws IOException
	 */
	public static ExcelParseFunctions initByInputStream(InputStream inputStream) throws IOException {
		return new ExcelParseFunctions().init(inputStream);
	}

	public ExcelParseFunctions init(InputStream inputStream) throws IOException {
		this.workbook = ExcelUtil.parseInputStreamToWorkBook(inputStream);
		return this;
	}

	public ExcelParseFunctions init(File file) throws IOException {
		this.workbook = ExcelUtil.parseFileToWorkBook(file);
		return this;
	}

	public ExcelParseFunctions addSheetRules(SheetParseRule parseSheetRule) {
		sheetRules.add(parseSheetRule);
		return this;
	}

	public ExcelParseFunctions addSheetRules(List<SheetParseRule> parseSheetRule) {
		sheetRules.addAll(parseSheetRule);
		return this;
	}

	/**
	 * @param parseSheetRule
	 * @return
	 */
	public ExcelParseFunctions addSheetRules(SheetParseRule... parseSheetRule) {
		for (SheetParseRule sheetParseRule : parseSheetRule) {
			this.addSheetRules(sheetParseRule);
		}
		return this;
	}


	/**
	 * 调用前必须设置SheetRules  参照以下方法
	 * <p>{@link ExcelParseFunctions#addSheetRules(SheetParseRule)}</p>
	 * <p>{@link ExcelParseFunctions#addSheetRules(SheetParseRule...)}</p>
	 * <p>{@link ExcelParseFunctions#addSheetRules(List)}</p>
	 * 按照 sheetParseRule进行解析
	 */
	public ExcelParseResultModel parse() {
		ExcelParseResultModel excelParseResultModel = new ExcelParseResultModel();
		MutilsException.throwException(sheetRules.isEmpty(), "必须要填写sheet解析规则");
		for (SheetParseRule rows : sheetRules) {
			int sheetIndex = rows.getSheetIndex();
			Sheet aimSheet = workbook.getSheetAt(sheetIndex);
			ExcelSheetModel excelSheetModel = new ExcelSheetModel();
			excelSheetModel.setSheetName(aimSheet.getSheetName());
			excelSheetModel.setSheetIndex(sheetIndex);
			int lastRowNum = aimSheet.getLastRowNum();

			for (int i = rows.getParseRowRule().getStartRowIndex(); i <= lastRowNum; i++) {
				Row next = aimSheet.getRow(i);
				if (next == null) {
					continue;
				}
				ExcelRowModel excelRowModel = new ExcelRowModel();
				excelRowModel.setRowIndex(next.getRowNum());
				for (int needCell : rows.getParseRowRule().getParseCellIndexes()) {
					Cell cell = next.getCell(needCell);
					if (cell != null) {
						excelRowModel.addCells(new ExcelCellModel(needCell, cell));
					}
				}
				excelSheetModel.addRows(excelRowModel);
			}
			excelParseResultModel.addSheet(excelSheetModel);
		}
		return excelParseResultModel;
	}
}

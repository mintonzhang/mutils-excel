package cn.minsin.excel.function;

import cn.minsin.excel.enums.ExcelVersion;
import cn.minsin.excel.function.creator.BaseCreator;
import cn.minsin.excel.function.creator.GetWorkbook;
import cn.minsin.excel.function.creator.SheetCreator;
import cn.minsin.excel.tools.ExcelUtil;
import lombok.Getter;
import lombok.NonNull;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import java.io.IOException;
import java.io.InputStream;
import java.util.function.Consumer;

/**
 * @author: minton.zhang
 * @since: 2020/5/5 14:13
 */
@Getter
@Setter
public class ExcelCreateFunctions extends BaseCreator implements GetWorkbook {

	private ExcelCreateFunctions(Workbook workbook, ExcelVersion excelVersion) {
		super(workbook, excelVersion);
	}

	public ExcelCreateFunctions() {
		super(null, null);
	}

	public static ExcelCreateFunctions initByVersion(ExcelVersion excelVersion) {
		return new ExcelCreateFunctions().init(excelVersion);
	}

	public static ExcelCreateFunctions initByInputStream(InputStream in) throws IOException {
		return new ExcelCreateFunctions().init(in);
	}

	public static ExcelCreateFunctions initByTemplatePath(String templatePath, boolean isInDisk) throws IOException {
		return new ExcelCreateFunctions().init(templatePath, isInDisk);
	}


	public ExcelCreateFunctions init(@NonNull ExcelVersion excelVersion) {
		this.workbook = excelVersion.createInstance();
		this.excelVersion = excelVersion;
		return this;
	}

	public ExcelCreateFunctions init(InputStream in) throws IOException {
		this.workbook = WorkbookFactory.create(in);
		this.excelVersion = ExcelVersion.checkVersion(workbook);
		return this;
	}

	public ExcelCreateFunctions init(String templatePath, boolean isInDisk) throws IOException {
		InputStream excelTemplate = ExcelUtil.getExcelTemplate(templatePath, isInDisk);
		return init(excelTemplate);
	}


	/**
	 * 创建sheet如果sheet存在 则读取
	 *
	 * @param index sheet下标
	 * @param name  sheet 名称
	 * @return
	 */
	public SheetCreator createSheet(int index, String name) {
		return this.createSheet(index, name, 0);
	}

	/**
	 * 创建sheet如果sheet存在 则读取
	 *
	 * @param index         sheet下标
	 * @param name          sheet 名称
	 * @param startRowIndex 写入开始开行下标 之后将会以此行开始计算
	 * @return
	 */
	public SheetCreator createSheet(int index, String name, int startRowIndex) {
		Sheet sheet;
		try {
			sheet = workbook.getSheetAt(index);
		} catch (Exception e) {
			sheet = workbook.createSheet(name);
		}
		return new SheetCreator(sheet, this, startRowIndex);
	}

	/**
	 * 消费函数
	 */
	public ExcelCreateFunctions workbookConsumer(Consumer<Workbook> consumer) {
		consumer.accept(this.workbook);
		return this;
	}

	/**
	 * 获取版本后缀
	 */
	@Override
	public String getVersionSuffix() {
		return this.excelVersion.getSuffix();
	}


}

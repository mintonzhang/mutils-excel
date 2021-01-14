package cn.minsin.excel.function.creator;

import cn.minsin.core.exception.MutilsException;
import cn.minsin.excel.function.ExcelCreateFunctions;
import cn.minsin.excel.model.create.ExcelExportTemplate;
import cn.minsin.excel.model.create.FiledParse;
import cn.minsin.excel.tools.ExcelUtil;
import lombok.AccessLevel;
import lombok.Getter;
import lombok.Setter;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.List;
import java.util.concurrent.ConcurrentHashMap;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.function.Consumer;

/**
 * @author: minton.zhang
 * @since: 2020/5/5 14:12
 */
@Getter
@Setter
public class SheetCreator extends BaseCreator implements GetWorkbook {
	@Setter(AccessLevel.NONE)
	private final ConcurrentHashMap<String, Integer> mergedRegion = new ConcurrentHashMap<>(10);
	private final Sheet sheet;
	private final ExcelCreateFunctions excelCreateFunctions;
	/**
	 * 默认写入开始行下标
	 */
	@Setter(AccessLevel.NONE)
	private final AtomicInteger startRowIndex;


	//********************************分割线****************************************//

	//以下方法 会根据创建sheet时 设置的默认开始行 自动增长行

	public SheetCreator(Sheet sheet, ExcelCreateFunctions excelCreateFunctions, int startRowIndex) {
		super(excelCreateFunctions.workbook, excelCreateFunctions.excelVersion);
		this.sheet = sheet;
		this.excelCreateFunctions = excelCreateFunctions;
		this.startRowIndex = new AtomicInteger(startRowIndex);
	}

	public RowCellCreator creatorRowWithAutoGrow() {
		Row row = sheet.createRow(startRowIndex.addAndGet(1));
		return new RowCellCreator(row, this);
	}

	public RowCellCreator cloneRowWithAutoGrow(int targetIndex, boolean afterRemove) {
		Row row = sheet.getRow(targetIndex);
		if (afterRemove) {
			sheet.removeRow(row);
		}
		Row newRow = sheet.createRow(startRowIndex.addAndGet(1));
		ExcelUtil.copyRow(row, newRow, true);
		return new RowCellCreator(row, this);
	}

	/**
	 * 按照list的规则进行生成excel
	 * <p>需要使用{@link cn.minsin.excel.model.create.CellTitle}</p>
	 * <p>需要使用{@link ExcelExportTemplate}</p>
	 *
	 * @return
	 */
	public SheetCreator addRowByListWithAutoGrow(List<? extends ExcelExportTemplate> listData) throws IllegalAccessException {
		return addRowByListWithAutoGrow(listData, false);
	}

	/**
	 * 按照list的规则进行生成excel
	 * <p>需要使用{@link cn.minsin.excel.model.create.CellTitle}</p>
	 * <p>需要使用{@link ExcelExportTemplate}</p>
	 *
	 * @return
	 */
	public SheetCreator addRowByListWithAutoGrow(List<? extends ExcelExportTemplate> listData, boolean needTitle) throws IllegalAccessException {
		MutilsException.throwException(CollectionUtils.isEmpty(listData), "需要生成Excel的数据不能为空");
		int title = 0;
		if (needTitle) {
			title = 1;
			List<FiledParse> parse = listData.get(0).parse();
			//创建行
			RowCellCreator rowCellCreator = this.creatorRowWithAutoGrow();
			for (int i = 0; i < parse.size(); i++) {
				//创建列
				rowCellCreator.cell(i, CellType.STRING, parse.get(i).getCellTitle().value());
			}
		}
		//创建excel
		for (int i = 0; i < listData.size(); i++) {
			listData.get(i).create(this, i + startRowIndex.addAndGet(1) + title);
		}
		return this;
	}

	//********************************分割线****************************************//


	public RowCellCreator creatorRow(int index) {
		Row row = sheet.createRow(index);
		return new RowCellCreator(row, this);
	}

	public RowCellCreator getRow(int index, boolean afterRemove) {
		Row row = sheet.getRow(index);
		if (afterRemove) {
			sheet.removeRow(row);
		}
		return new RowCellCreator(row, this);
	}

	public RowCellCreator cloneRow(int targetIndex, int newIndex, boolean afterRemove) {
		Row row = sheet.getRow(targetIndex);
		if (afterRemove) {
			sheet.removeRow(row);
		}
		Row newRow = sheet.createRow(newIndex);
		ExcelUtil.copyRow(row, newRow, true);
		return new RowCellCreator(row, this);
	}


	/**
	 * 合并单元格
	 *
	 * @param startRowIndex
	 * @param endRowIndex
	 * @param startColIndex
	 * @param endColIndex
	 * @param isSafe
	 * @return
	 */
	public SheetCreator addMergedRegion(String name, int startRowIndex, int endRowIndex, int startColIndex, int endColIndex, boolean isSafe) {
		//判断是否存在
		MutilsException.throwException(this.mergedRegion.containsKey(name), "存在相同的key单元格名称'" + name + "'");

		CellRangeAddress cellAddresses = new CellRangeAddress(startRowIndex, endRowIndex, startColIndex, endColIndex);

		int i = isSafe ? this.sheet.addMergedRegion(cellAddresses) : this.sheet.addMergedRegionUnsafe(cellAddresses);
		this.mergedRegion.put(name, i);
		return this;
	}

	/**
	 * 合并单元格
	 *
	 * @param startRowIndex
	 * @param endRowIndex
	 * @param startColIndex
	 * @param endColIndex
	 * @return
	 */
	public SheetCreator addMergedRegion(String name, int startRowIndex, int endRowIndex, int startColIndex, int endColIndex) {
		//判断是否存在
		this.addMergedRegion(name, startRowIndex, endRowIndex, startColIndex, endColIndex, false);
		return this;
	}

	/**
	 * 移除合并单单元格
	 *
	 * @param name
	 * @return
	 */
	public SheetCreator removeMergedRegion(String name) {
		this.mergedRegion.computeIfPresent(name, (k, v) -> {
			this.sheet.removeMergedRegion(v);
			this.mergedRegion.remove(name);
			return null;
		});
		return this;
	}

	/**
	 * 消费函数
	 *
	 * @param consumer
	 * @return
	 */
	public SheetCreator sheetConsumer(Consumer<Sheet> consumer) {
		consumer.accept(this.sheet);
		return this;
	}

	/**
	 * 按照list的规则进行生成excel
	 * <p>需要使用{@link cn.minsin.excel.model.create.CellTitle}</p>
	 * <p>需要使用{@link ExcelExportTemplate}</p>
	 *
	 * @return
	 */
	public SheetCreator addRowByList(List<? extends ExcelExportTemplate> listData) throws IllegalAccessException {
		return addRowByList(listData, 0, false);
	}

	/**
	 * 按照list的规则进行生成excel
	 * <p>需要使用{@link cn.minsin.excel.model.create.CellTitle}</p>
	 * <p>需要使用{@link ExcelExportTemplate}</p>
	 *
	 * @return
	 */
	public SheetCreator addRowByList(List<? extends ExcelExportTemplate> listData, int startRowIndex, boolean needTitle) throws IllegalAccessException {
		MutilsException.throwException(CollectionUtils.isEmpty(listData), "需要生成Excel的数据不能为空");
		int title = 0;
		if (needTitle) {
			title = 1;
			List<FiledParse> parse = listData.get(0).parse();
			//创建行
			RowCellCreator rowCellCreator = this.creatorRow(startRowIndex);
			for (int i = 0; i < parse.size(); i++) {
				//创建列
				rowCellCreator.cell(i, CellType.STRING, parse.get(i).getCellTitle().value());
			}
		}
		//创建excel
		for (int i = 0; i < listData.size(); i++) {
			listData.get(i).create(this, i + startRowIndex + title);
		}
		return this;
	}

	@Override
	public Workbook getWorkbook() {
		return sheet.getWorkbook();
	}
}

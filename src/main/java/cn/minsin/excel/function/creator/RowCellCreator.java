package cn.minsin.excel.function.creator;

import cn.minsin.excel.tools.ExcelUtil;
import lombok.Getter;
import lombok.NonNull;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.function.Consumer;

/**
 * @author: minton.zhang
 * @since: 2020/5/5 14:12
 */
@Getter
public class RowCellCreator extends BaseCreator {

	private final Row row;

	private final SheetCreator sheetCreator;

	public RowCellCreator(Row row, SheetCreator sheetCreator) {
		super(sheetCreator.workbook, sheetCreator.excelVersion);
		this.row = row;
		this.sheetCreator = sheetCreator;
	}

	public RowCellCreator cell(int index, CellType cellType, Object value) {
		return this.cell(index, cellType, value, null);
	}

	public RowCellCreator cell(int index, CellType cellType, Object value, Consumer<CellStyle> cellStyle) {
		ExcelUtil.setCellValue(this.row, index, cellType, value, cellStyle);
		return this;
	}

	public RowCellCreator removeCell(int index) {
		this.removeCell(index);
		return this;
	}

	/**
	 * 设置cellStyle
	 *
	 * @param collIndex         列下标
	 * @param cellStyleFunction 列样式 消费函数
	 */
	public RowCellCreator setCellStyle(int collIndex, Consumer<CellStyle> cellStyleFunction) {
		@NonNull
		Cell cell = row.getCell(collIndex);
		CellStyle cellStyle = this.createCellStyle();
		cellStyleFunction.accept(cellStyle);
		cell.setCellStyle(cellStyle);
		return this;
	}

	/**
	 * 设置cellStyle
	 *
	 * @param cellStyleFunction 列样式 消费函数
	 */
	public RowCellCreator setRowStyle(Consumer<CellStyle> cellStyleFunction) {
		CellStyle cellStyle = this.createCellStyle();
		cellStyleFunction.accept(cellStyle);
		row.setRowStyle(cellStyle);
		return this;
	}


	/**
	 * 消费函数
	 *
	 * @param consumer
	 * @return
	 */
	public RowCellCreator cell(Consumer<Row> consumer) {
		consumer.accept(this.row);
		return this;
	}


	@Override
	public Workbook getWorkbook() {
		return row.getSheet().getWorkbook();
	}

	/**
	 * 创建cellStyle
	 *
	 * @return
	 */
	public CellStyle createCellStyle() {
		return this.getWorkbook().createCellStyle();
	}
}

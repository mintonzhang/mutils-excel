package cn.minsin.excel.model.create;

import cn.minsin.excel.function.creator.RowCellCreator;
import cn.minsin.excel.function.creator.SheetCreator;
import org.apache.poi.ss.usermodel.CellType;

import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

/**
 * @author: minton.zhang
 * @since: 2020/4/8 23:28
 */
public interface ExcelExportTemplate {

	default void create(SheetCreator sheetCreator, int rowIndex) throws IllegalAccessException {
		RowCellCreator rowCellCreator = sheetCreator.creatorRow(rowIndex);
		List<FiledParse> parse = this.parse();
		for (int i = 0; i < parse.size(); i++) {
			rowCellCreator.cell(i, CellType.STRING, parse.get(i).getValue());
		}
	}


	default List<FiledParse> parse() throws IllegalAccessException {

		Field[] declaredFields = this.getClass().getDeclaredFields();
		ArrayList<FiledParse> filedParses = new ArrayList<>(declaredFields.length);
		for (Field declaredField : declaredFields) {
			CellTitle annotation = declaredField.getAnnotation(CellTitle.class);
			if (annotation == null) {
				continue;
			}
			declaredField.setAccessible(true);
			Object o = declaredField.get(this);
			filedParses.add(new FiledParse(annotation, o));
		}
		filedParses.trimToSize();
		return filedParses;
	}

}

package cn.minsin.excel.function.creator;

import cn.minsin.core.tools.IOUtil;
import cn.minsin.excel.enums.ExcelVersion;
import cn.minsin.excel.tools.ExcelUtil;
import lombok.AllArgsConstructor;
import lombok.NonNull;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Workbook;

import javax.servlet.http.HttpServletResponse;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.function.Consumer;
import java.util.function.Function;

/**
 * @author: minton.zhang
 * @since: 2020/5/11 21:27
 */
@AllArgsConstructor
public class BaseCreator implements GetWorkbook {

	@Setter
	protected Workbook workbook;

	@Setter
	protected ExcelVersion excelVersion;

	public String getVersionSuffix() {
		return excelVersion.getSuffix();
	}

	/**
	 * 导出到文件流
	 *
	 * @param outputStream 输出流
	 * @throws IOException
	 */
	public void export(OutputStream outputStream) throws IOException {
		try {
			workbook.write(outputStream);
		} finally {
			IOUtil.close(workbook, outputStream);
		}
	}

	/**
	 * 导出文件到浏览器
	 *
	 * @param resp
	 * @param fileName 无后缀的文件名 2018年10月11日
	 */
	public void export(HttpServletResponse resp, String fileName) throws Exception {
		ExcelUtil.workbookToServletResponse(this.workbook, resp, fileName);
	}

	/**
	 * 导出文件到浏览器 并且设置文件名编码格式
	 *
	 * @param resp
	 * @param fileName 无后缀的文件名 2018年10月11日
	 */
	public void exportWithUTF8(HttpServletResponse resp, String fileName) throws Exception {
		ExcelUtil.workbookToServletResponse(this.workbook, resp, new String(fileName.getBytes(StandardCharsets.UTF_8), StandardCharsets.ISO_8859_1));
	}


	/**
	 * 将workbook转换成为InputStream 并且提供转换转换函数
	 * 通常用于将workbook上传到第三方oss
	 *
	 * @param convert 转换函数
	 * @param <T>
	 * @return
	 * @throws Exception
	 */
	public <T> T toInputStream(@NonNull final Function<InputStream, T> convert) throws Exception {
		return ExcelUtil.workbookToInputStream(this.workbook, convert);
	}

	/**
	 * 将workbook转换成InputStream
	 */
	public InputStream toInputStream() throws Exception {
		return ExcelUtil.workbookToInputStream(this.workbook, e -> e);
	}

	/**
	 * 将workbook转换成为InputStream 并且提供转换转换函数
	 * 通常用于将workbook进行消费 比如使用{@link HttpServletResponse}
	 *
	 * @param convert 转换函数
	 * @throws Exception
	 */
	public void toInputStream(@NonNull final Consumer<InputStream> convert) throws Exception {
		ExcelUtil.workbookToInputStream(this.workbook, convert);
	}

	/**
	 * 将workbook转换成为InputStream 并且提供转换转换函数
	 * 通常用于将workbook进行消费 比如使用{@link HttpServletResponse}
	 *
	 * @param convert 转换函数
	 * @throws Exception
	 */
	public <T> T uploadToOss(@NonNull final Function<InputStream, T> convert) throws Exception {
		return ExcelUtil.workbookToInputStream(this.workbook, convert);
	}

	@Override
	public Workbook getWorkbook() {
		return this.workbook;
	}
}

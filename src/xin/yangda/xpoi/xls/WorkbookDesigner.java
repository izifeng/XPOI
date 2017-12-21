package xin.yangda.xpoi.xls;

import java.io.BufferedInputStream;
import java.io.BufferedOutputStream;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;

import com.alibaba.fastjson.JSONObject;

import xin.yangda.xpoi.xls.entity.DataSourceEntity;
import xin.yangda.xpoi.xls.entity.DataSourceFieldEntity;
import xin.yangda.xpoi.xls.util.WordDefined;

/**
 * 根据特定Excel导出Excel文档
 * 
 * @author izifeng
 * @version 2.0
 * @date 2017-12-12 14:03
 * @site https://gitee.com/izifeng/XPOI.git
 *
 */
abstract class WorkbookDesigner {

	/**
	 * 工作表
	 */
	protected HSSFWorkbook workBook;

	/**
	 * 文档输出流
	 */
	protected ByteArrayOutputStream outStream;

	/**
	 * 单个数据源
	 */
	protected DataSourceEntity dataSource;

	/**
	 * 多个数据源
	 */
	protected List<DataSourceEntity> dataSourceList;

	public WorkbookDesigner() {
		outStream = new ByteArrayOutputStream();
	}

	public ByteArrayOutputStream getOutStream() {
		return outStream;
	}

	/**
	 * 打开excel模板
	 * 
	 * @param filePath
	 *            excel模板路径
	 * @throws IOException
	 */
	public void open(String filePath) throws IOException {
		if (filePath.trim().toLowerCase().endsWith("xls")) {
			File file = new File(filePath);

			FileInputStream fis = new FileInputStream(file);

			POIFSFileSystem fs = new POIFSFileSystem(fis);

			// 读取excel模板
			workBook = new HSSFWorkbook(fs);
		} else {
			throw new IllegalArgumentException(WordDefined.ONLY_SUPPORT_XLS);
		}
	}

	/**
	 * 打开excel模板
	 * 
	 * @param filePath
	 *            excel模板路径
	 * @throws IOException
	 */
	public void open(InputStream inputStream) throws IOException {
		POIFSFileSystem fs = new POIFSFileSystem(inputStream);
		// 读取excel模板
		workBook = new HSSFWorkbook(fs);
	}

	/**
	 * 添加数据源
	 * 
	 * @param dataSource
	 */
	public void setDataSource(DataSourceEntity dataSource) {
		this.dataSource = dataSource;
	}

	/**
	 * 添加数据源
	 * 
	 * @param dataSource
	 */
	public void setDataSource(List<DataSourceEntity> dataSourceList) {
		this.dataSourceList = dataSourceList;
	}

	/**
	 * 数据处理
	 * 
	 * @throws IOException
	 */
	public abstract void process() throws IOException;

	/**
	 * 保存文件
	 * 
	 * @param filePath
	 *            目标文件路径
	 * @throws IOException
	 */
	public void saveFile(String filePath) throws IOException {
		// 获取流字节
		byte[] content = outStream.toByteArray();

		// 将流字节转换成输入流
		InputStream is = new ByteArrayInputStream(content);

		// 获取目标文件输出流
		FileOutputStream outputStream = new FileOutputStream(filePath);

		// 获取缓冲字节输入流
		BufferedInputStream bis = new BufferedInputStream(is);

		// 获取缓冲字节输出流
		BufferedOutputStream bos = new BufferedOutputStream(outputStream);

		byte[] buff = new byte[8192];
		int bytesRead;
		while (-1 != (bytesRead = bis.read(buff, 0, buff.length))) {
			bos.write(buff, 0, bytesRead);
		}

		bis.close();
		bos.close();
		outputStream.flush();
		outputStream.close();
	}

	/**
	 * 写自定义字段试数据
	 * 
	 * @param sheet
	 *            工作簿
	 * @param cdsField
	 *            工作簿中自定义字段位置信息
	 * @param dataSource
	 *            数据源
	 */
	public void writeCustomField(HSSFSheet sheet, Map<String, DataSourceFieldEntity> cdsField,
			DataSourceEntity dataSource) {
		Map<String, String> customField = dataSource.getCustomField();
		for (Map.Entry<String, DataSourceFieldEntity> entry : cdsField.entrySet()) {
			String key = entry.getKey();
			DataSourceFieldEntity dsf = entry.getValue();

			if (!customField.containsKey(key))
				continue;

			// 获取行
			Row row = sheet.getRow(dsf.getRowIdx());
			// 获取列
			Cell cell = row.getCell(dsf.getColIdx());
			// 填值
			cell.setCellValue(customField.get(key));
		}
	}

	/**
	 * 写数据源字段数据
	 * 
	 * @param sheet
	 *            工作簿
	 * @param dsField
	 *            工作簿中数据源字段位置信息
	 * @param dataSource
	 *            数据源
	 */
	public void writeDataSource(HSSFSheet sheet, Map<String, DataSourceFieldEntity> dsField,
			DataSourceEntity dataSource) {
		List<JSONObject> ds = dataSource.getDataSource();
		for (Map.Entry<String, DataSourceFieldEntity> entry : dsField.entrySet()) {
			String key = entry.getKey();
			DataSourceFieldEntity sdf = entry.getValue();
			for (int i = 0; i < ds.size(); i++) {
				JSONObject json = ds.get(i);
				if (!json.containsKey(key))
					continue;

				// 获取行
				Row row = sheet.getRow(sdf.getRowIdx() + i);
				if (row == null)
					row = sheet.createRow(sdf.getRowIdx() + i);

				// 获取列
				Cell cell = row.getCell(sdf.getColIdx());
				if (cell == null)
					cell = row.createCell(sdf.getColIdx());

				// 设置单元格样式
				cell.setCellStyle(sdf.getCellStyle());

				// 填充单元格数据
				cell.setCellValue(json.getString(key));
			}
		}
	}

	/**
	 * 获取工作簿
	 * 
	 * @param index
	 *            工作簿索引
	 * @return
	 */
	protected HSSFSheet getHSSFSheet(int index) {
		return workBook.getSheetAt(index);
	}

	/**
	 * 提取模板中数据源相关字段位置信息，包括：行、列、单元格样式
	 * 
	 * @param sheet
	 *            工作簿
	 * @param map1
	 *            自定义字段集
	 * @param map2
	 *            数据源字段集
	 */
	public void getDataSourceField(HSSFSheet sheet, Map<String, DataSourceFieldEntity> map1,
			Map<String, DataSourceFieldEntity> map2) {
		// 获得总列数
		int coloumNum = sheet.getRow(0).getPhysicalNumberOfCells();

		// 获得总行数
		int rowNum = sheet.getLastRowNum();

		// 提取模板中数据源相关字段位置信息，包括：行、列、单元格样式
		for (int rowIdx = 0; rowIdx <= rowNum; rowIdx++) {
			Row row = sheet.getRow(rowIdx);
			for (int colIdx = 0; colIdx < coloumNum; colIdx++) {
				// 获取单元格
				Cell cell = row.getCell(colIdx);

				// 单元格值
				String cellValue = cell.getStringCellValue().trim();

				// 单元格内容不满足模板要求
				if (cellValue.length() == 0 || cellValue.indexOf("&=") != 0)
					continue;

				DataSourceFieldEntity dsf = new DataSourceFieldEntity(rowIdx, colIdx, cell.getCellStyle());

				if (cellValue.indexOf("&=$") == 0) {// 自定义数据源字段
					String field = cellValue.replace("&=$", "");
					map1.put(field, dsf);
				} else if (cellValue.indexOf("&=") == 0) {// 数据源字段
					String field = cellValue.replace("&=", "");
					map2.put(field, dsf);
				}
			}
		}
	}
}

package xin.yangda.xpoi.xls;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFSheet;

import xin.yangda.xpoi.xls.entity.DataSourceFieldEntity;

/**
 * 根据特定Excel导出Excel文档（支持单个或多个工作簿的excel模板）
 * 
 * @author izifeng
 * @version 2.0
 * @date 2017-12-12 14:03
 * @site https://gitee.com/izifeng/XPOI.git
 *
 */
public class MultipleWorkbookDesigner extends WorkbookDesigner {

	@Override
	public void process() throws IOException {
		// 获取工作簿页数
		int sheetCount = this.workBook.getNumberOfSheets();
		for (int i = 0; i < sheetCount; i++) {
			HSSFSheet sheet = this.getHSSFSheet(i);

			// 工作簿自定义字段集
			Map<String, DataSourceFieldEntity> cdsField = new HashMap<>();

			// 工作簿数据源字段集
			Map<String, DataSourceFieldEntity> dsField = new HashMap<>();

			// 读取工作簿自定义字段，数据源字段数据集
			this.getDataSourceField(sheet, cdsField, dsField);

			// 写自定义字段
			if (cdsField.size() > 0) {
				this.writeCustomField(sheet, cdsField, dataSourceList.get(i));
			}

			// 写数据源
			if (dsField.size() > 0) {
				this.writeDataSource(sheet, dsField, dataSourceList.get(i));
			}
		}

		this.workBook.write(outStream);
		this.workBook.close();
	}

}

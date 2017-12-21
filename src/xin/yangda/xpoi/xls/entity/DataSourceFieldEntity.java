package xin.yangda.xpoi.xls.entity;

import org.apache.poi.ss.usermodel.CellStyle;

/**
 * 模板数据源字段位置信息实体
 * 
 * @author izifeng
 * @version 2.0
 * @date 2017-12-12 10:44
 * @site https://gitee.com/izifeng/XPOI.git
 *
 */
public class DataSourceFieldEntity {

	private int rowIdx;

	private int colIdx;

	private CellStyle cellStyle;

	public DataSourceFieldEntity() {
		// TODO Don't do anything
	}

	public DataSourceFieldEntity(int rowIdx, int colIdx, CellStyle cellStyle) {
		this.rowIdx = rowIdx;
		this.colIdx = colIdx;
		this.cellStyle = cellStyle;
	}

	public int getRowIdx() {
		return rowIdx;
	}

	public void setRowIdx(int rowIdx) {
		this.rowIdx = rowIdx;
	}

	public int getColIdx() {
		return colIdx;
	}

	public void setColIdx(int colIdx) {
		this.colIdx = colIdx;
	}

	public CellStyle getCellStyle() {
		return cellStyle;
	}

	public void setCellStyle(CellStyle cellStyle) {
		this.cellStyle = cellStyle;
	}

}

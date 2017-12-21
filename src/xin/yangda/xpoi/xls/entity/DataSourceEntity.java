package xin.yangda.xpoi.xls.entity;

import java.util.List;
import java.util.Map;

import com.alibaba.fastjson.JSONObject;

/**
 * 数据源实体
 * 
 * @author izifeng
 * @version 2.0
 * @date 2017-12-12 10:44
 * @site https://gitee.com/izifeng/XPOI.git
 *
 */
public class DataSourceEntity {

	/**
	 * 自定义字段
	 */
	private Map<String, String> customField;

	/**
	 * 数据源
	 */
	private List<JSONObject> dataSource;

	public DataSourceEntity() {
		// TODO Don't do anything
	}

	public DataSourceEntity(Map<String, String> customField, List<JSONObject> dataSource) {
		this.customField = customField;
		this.dataSource = dataSource;
	}

	public Map<String, String> getCustomField() {
		return customField;
	}

	public void setCustomField(Map<String, String> customField) {
		this.customField = customField;
	}

	public List<JSONObject> getDataSource() {
		return dataSource;
	}

	public void setDataSource(List<JSONObject> dataSource) {
		this.dataSource = dataSource;
	}
}

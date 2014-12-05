package jenfong.hnacode.deal;

import java.io.FileInputStream;
import java.io.IOException;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import net.sf.jxls.exception.ParsePropertyException;
import net.sf.jxls.transformer.XLSTransformer;

import org.apache.commons.lang.StringUtils;
import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Excelsplit {

	private static Log log = LogFactory.getLog(Excelsplit.class);
	private static String templateFileName = "D://edata/dealfiles/dailyReportTemplate_dev2.xls";
	// private static String templateFileName =
	// "D://edata/dealfiles/dailyReportTemplate.xls";
	private static String destFileName = "D://edata/dealfiles/dest_test/康乐园2014-11-";
	private static String destFileName1 = "D://edata/dealfiles/dest_test/康乐园2014-10-";

	public static void main(String[] args) {
		// String sourcefile = "D:/edata/Book1.xls";
		String sourcefile = "D:/edata/康乐园2014-11-25.xls";
		try {
			Workbook wb = createWb(sourcefile);
			HashMap<String, Object> sheetMap = new HashMap<String, Object>();
			int sheetNum = wb.getNumberOfSheets();
			Sheet sheet;
			String sheetName;
			DaliyDataBean ddBean;
			for (int i = 0; i < sheetNum; i++) {
				// sheetMap.put(sheetName,null);
				sheet = wb.getSheetAt(i);
				sheetName = sheet.getSheetName();
				ddBean = new DaliyDataBean();
				log.debug(sheetName);
				int rowNum = sheet.getPhysicalNumberOfRows();
				// int rowNum = sheet.getLastRowNum();
				Row row;
				List<Object> pickupData = new ArrayList<Object>();
				for (int j = 0; j < rowNum; j++) {
					if (j < 13 || j > 32)
						continue;
					row = sheet.getRow(j);
					// log.info(name.getSheetName()+"_"+getValueFromCell(row.getCell(3)));
					log.debug("[表名:" + sheetName + "_" + (j + 1) + "行_3列] "
							+ getValueFromCell(row.getCell(3)));
					pickupData.add(getValueFromCell(row.getCell(3)));
					if (j == 30)
						pickupData.add(new Double(0));// 第24行，缺少营业外
				}
				// log.info(sheet.getPhysicalNumberOfRows());

				// BeanUtils.copyProperties(ddBean, pickupData);
				if (pickupData.isEmpty())
					continue;
				// 封装获取的原数据
				Class userCla = (Class) ddBean.getClass();
				Field[] fs = userCla.getDeclaredFields();

				for (int m = 0; m < fs.length; m++) {
					Field f = fs[m];
					f.setAccessible(true);
					try {
						f.set(ddBean, pickupData.get(m));
					} catch (IllegalArgumentException e) {
						e.printStackTrace();
					} catch (IllegalAccessException e) {
						e.printStackTrace();
					}
				}

				// sheetMap.put(sheetName, ddBean);
				// if(sheetName.equals("1")){
				try {
					Integer sheetNameNum = Integer.valueOf(sheetName);
					String destFileNameAppend = destFileName;
					if (sheetNameNum > 25)// 财务月
						destFileNameAppend = destFileName1;

					if (sheetName.length() == 1)
						sheetName = "0" + sheetName;// 不足两位，前置补0
					destFileNameAppend = destFileNameAppend + sheetName
							+ ".xls";// 2014-10-26~
					// log.info(destFileNameAppend);
					String[] nameStrArr = destFileNameAppend.split("/");
					log.info("call D:\\BI\\kettle\\Kitchen.bat /rep=hotel /job job/job_budget_daily_oq_excel -param filename=D:\\excel\\"
							+ nameStrArr[nameStrArr.length - 1]
							+ " >>D:\\excel\\aa.log");
					exportExcel(templateFileName, ddBean, destFileNameAppend);
				} catch (Exception e) {
				}
				// }
			}
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public static final Workbook createWb(String filePath) throws IOException {
		if (StringUtils.isBlank(filePath)) {
			throw new IllegalArgumentException("参数错误!!!");
		}
		if (filePath.trim().toLowerCase().endsWith("xls")) {
			return new HSSFWorkbook(new FileInputStream(filePath));
		} else if (filePath.trim().toLowerCase().endsWith("xlsx")) {
			return new XSSFWorkbook(new FileInputStream(filePath));
		} else {
			throw new IllegalArgumentException("不支持除：xls/xlsx以外的文件格式!!!");
		}
	}

	public static final Sheet getSheet(Workbook wb, String sheetName) {
		return wb.getSheet(sheetName);
	}

	public static final Sheet getSheet(Workbook wb, int index) {
		return wb.getSheetAt(index);
	}

	/**
	 * 获取单元格内文本信息
	 * 
	 * @param cell
	 * @return
	 * @date 2013-5-8
	 */
	public static final Object getValueFromCell(Cell cell) {
		if (cell == null) {
			log.debug("Cell is null !!!");
			return null;
		}
		Object value = null;
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_NUMERIC: // 数字
			if (HSSFDateUtil.isCellDateFormatted(cell)) { // 如果是日期类型
				value = new SimpleDateFormat("yyyy-MM-dd").format(cell
						.getDateCellValue());
			} else

				value = getRoundHalfUp(cell.getNumericCellValue());
			break;
		case Cell.CELL_TYPE_STRING: // 字符串
			value = cell.getStringCellValue();
			break;
		case Cell.CELL_TYPE_FORMULA: // 公式
			// 用数字方式获取公式结果，根据值判断是否为日期类型

			double numericValue = cell.getNumericCellValue();
			/*
			 * if(HSSFDateUtil.isValidExcelDate(numericValue)) { // 如果是日期类型
			 * value = new
			 * SimpleDateFormat("yyyy-MM-dd").format(cell.getDateCellValue()) ;
			 * } else
			 */
			value = getRoundHalfUp(numericValue);
			break;
		case Cell.CELL_TYPE_BLANK: // 空白
			// value = StringUtils.EMPTY;
			value = new Double(0);
			break;
		/*
		 * case Cell.CELL_TYPE_BOOLEAN: // Boolean value =
		 * String.valueOf(cell.getBooleanCellValue()); break;
		 */
		/*
		 * case Cell.CELL_TYPE_ERROR: // Error，返回错误码 value =
		 * String.valueOf(cell.getErrorCellValue()); break;
		 */
		default:
			// value = StringUtils.EMPTY;
			value = new Double(0);
			break;
		}
		// 使用[]记录坐标
		return value;// + "["+cell.getRowIndex()+","+cell.getColumnIndex()+"]" ;
	}

	/**
	 * @param cell
	 * @return
	 */
	public static final Cell getCloneCell(Cell cell) {
		//ToDo:封装获取的cell
		return cell;
	}
	
	/**
	 * @param sheet
	 * @return
	 */
	public static final List<Cell> getCellListFromSheet(Sheet sheet){
		//ToDo:从制定Sheet获取所需内容
		List<Cell> cellList = new ArrayList<Cell>();
		return cellList;
	} 

	public static void exportExcel(String templateFileName,
			DaliyDataBean daliyDataBean, String destFileName) {
		Map<String, Object> beans = new HashMap<String, Object>();
		beans.put("ddbean", daliyDataBean);
		XLSTransformer transformer = new XLSTransformer();
		try {
			transformer.transformXLS(templateFileName, beans, destFileName);
		} catch (ParsePropertyException e) {
			e.printStackTrace();
		} catch (InvalidFormatException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

	public static Double getRoundHalfUp(double dobuleData) {
		return dobuleData;
		// return new BigDecimal(dobuleData).setScale(2,
		// BigDecimal.ROUND_HALF_UP).doubleValue();
	}
}

package jenfong.hnacode.deal;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.commons.logging.Log;
import org.apache.commons.logging.LogFactory;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;

public class ExcelFreezePane {
	private static Log log = LogFactory.getLog(ExcelFreezePane.class);

	public static void main(String[] args) {

		if (args.length == 0)
			throw new RuntimeException(
					"参数输入错误！！！文件完整路径、冻结的行、冻结的列，参数以空格分开。如：\"E:/海航酒店集团酒店运营情况综合日报.xls\" 2 3");
		String filePath = args[0];
		// String filePath = "E:/海航酒店集团酒店运营情况综合日报.xls";
		File oldfile = new File(filePath);
		if (!oldfile.exists())
			throw new RuntimeException("文件路径错误或文件不存在，请检查!!!");
		File newfile = new File(filePath + "_bak");

		int freezeCol = args[1].isEmpty() ? 1 : Integer.parseInt(args[1]);
		int freezeRow = args[2].isEmpty() ? 1 : Integer.parseInt(args[2]);
		try {
			Workbook wbook = Excelsplit.createWb(filePath);
			Sheet sheet = wbook.getSheetAt(0);// Clone sheet
			log.info(sheet.getSheetName());
			Row row = sheet.getRow(freezeRow);
			Cell cell = row.getCell(freezeCol);
			log.info((cell.getRowIndex() + 1) + "行 "
					+ (cell.getColumnIndex() + 1) + "列 "
					+ Excelsplit.getValueFromCell(cell));

			FreezeUtil.freezeSpColumn(sheet, freezeCol, freezeRow);

			try {
				// Export

				oldfile.renameTo(newfile);
				FileOutputStream os = new FileOutputStream(filePath);
				wbook.write(os);
				os.close();
			} catch (Exception e) {
				newfile.renameTo(oldfile);
			}

		} catch (IOException e) {
			e.printStackTrace();
		}
	}
}

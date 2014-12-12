package jenfong.hnacode.deal;

import org.apache.poi.ss.usermodel.Sheet;

/**
 * @author jf.wu
 *
 */
public class FreezeUtil {

	/**
	 * @param sheet
	 * @param colSplit
	 * @param rowSplit
	 *            colSplit - Horizonatal position of split. rowSplit - Vertical
	 *            position of split.
	 */
	public static void freezeSpColumn(Sheet sheet, int colSplit, int rowSplit) {
		sheet.createFreezePane(colSplit, rowSplit);
	}

	/**
	 * @param sheet
	 * @param colSplit
	 * @param rowSplit
	 * @param leftmostColumn
	 * @param topRow
	 *            colSplit - Horizonatal position of split. rowSplit - Vertical
	 *            position of split. leftmostColumn - Left column visible in
	 *            right pane. topRow - Top row visible in bottom pane
	 */
	public static void freezeSpColumn(Sheet sheet, int colSplit, int rowSplit,
			int leftmostColumn, int topRow) {
		sheet.createFreezePane(colSplit, rowSplit, leftmostColumn, topRow);
	}
}

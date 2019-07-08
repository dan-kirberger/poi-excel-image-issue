package com.dankirberger.demo.poiexcelimageissue;

import org.apache.poi.ss.usermodel.ClientAnchor;

public class Main {

	public static void main(String[] args) {
		new ExcelGenerator(ClientAnchor.AnchorType.MOVE_AND_RESIZE).generateExampleWorksheet();
		new ExcelGenerator(ClientAnchor.AnchorType.MOVE_DONT_RESIZE).generateExampleWorksheet();
	}
}

package pkg;

import java.io.File;
import java.io.IOException;
import java.util.HashMap;
import java.util.Iterator;
import java.util.Locale;
import java.util.Map;
import java.util.Map.Entry;

import jxl.Workbook;
import jxl.WorkbookSettings;
import jxl.format.UnderlineStyle;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableFont;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

public class WriteExcel {

	private String inputFile;
	private HashMap<String, String> textFileMap = new HashMap<String, String>();

	private WritableCellFormat timesBoldUnderline;
	private WritableCellFormat times;

	public WriteExcel(HashMap<String, String> textFileMap, String inputFile) {
		this.textFileMap = textFileMap;
		this.inputFile = inputFile;
		times = new WritableCellFormat(new WritableFont(WritableFont.TIMES, 10));
		timesBoldUnderline = new WritableCellFormat(new WritableFont(
				WritableFont.TIMES, 10, WritableFont.BOLD, false,
				UnderlineStyle.SINGLE));
	}

	public void write() throws IOException, WriteException {

		File file = new File(inputFile);
		WorkbookSettings wbSettings = new WorkbookSettings();

		wbSettings.setLocale(new Locale("en", "EN"));

		WritableWorkbook workbook = Workbook.createWorkbook(file, wbSettings);
		workbook.createSheet("Report", 0);
		WritableSheet excelSheet = workbook.getSheet(0);

		createContent(excelSheet);

		workbook.write();
		workbook.close();
	}

	private void createContent(WritableSheet sheet) throws RowsExceededException, WriteException {
		
		addCaption(sheet, 2, 3, "Location");
		addCaption(sheet, 8, 3, "Content");
		
		int row = 5;
		
		Iterator<Entry<String, String>> it = textFileMap.entrySet().iterator();
		
		while(it.hasNext()) {
			Map.Entry entry = (Map.Entry) it.next();			
			addLabel(sheet, 2, row, entry.getKey().toString());
			addLabel(sheet, 8, row, entry.getValue().toString());
			row = row + 2;
		}

	}

	private void addCaption(WritableSheet sheet, int column, int row, String s)
			throws RowsExceededException, WriteException {
		Label label;
		label = new Label(column, row, s, timesBoldUnderline);
		sheet.addCell(label);
	}
	
	private void addLabel(WritableSheet sheet, int column, int row, String s)
		      throws WriteException, RowsExceededException {
		    Label label;
		    label = new Label(column, row, s, times);
		    sheet.addCell(label);
		  }
}

package excel;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ABC {
	public static void main(String[] args) {
		Workbook wb = null;
		InputStream is;
		try {
			is = new FileInputStream(args[0]); //檔案位置
			wb = new XSSFWorkbook(is);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		Sheet sheet = wb.getSheetAt(0);

		if (null == sheet) {
			System.out.println("解析Excel失敗");
			return;
		}

		int firstRowNum = sheet.getFirstRowNum();
		Row firstRow = sheet.getRow(firstRowNum);
		if (null == firstRow) {
			System.out.println("解析Excel失敗");
			return;
		}
		int lastCellNum = firstRow.getLastCellNum();
		List<String> planName = new ArrayList<String>();
		Cell cell;
		for(int cellNum = 5; cellNum < lastCellNum; cellNum++) {
			cell = firstRow.getCell(cellNum);
			if (null == cell) {
				planName.add("");
			} else {
				planName.add(getCellValue(cell));
			}
		}
		
		List<List<String>> dataList = new ArrayList<List<String>>();
		
		String benefit = "";
		String Coverage = "";
		String Category = "";
		Boolean doContinue = false;
		
		int rowStart = firstRowNum + 1;
		int rowEnd = sheet.getLastRowNum();
		Row row;
		for (int rowNum = rowStart; rowNum <= rowEnd; rowNum++) {
			row = sheet.getRow(rowNum);
			if (null == row) {
				continue;
			}
			if(row.getZeroHeight()) {
				continue;
			}
			
			XSSFColor color;
			List<String> data;
			for(int cellNum = row.getFirstCellNum(); cellNum < lastCellNum; cellNum++) {
				cell = row.getCell(cellNum);
				if (null == cell) {
					continue;
				} else {
					color = (XSSFColor) cell.getCellStyle().getFillForegroundColorColor();
					if (null != color) {
						benefit = getCellValue(cell);
						break;
					}
					if(cellNum == 0) {
						if(!"".equals(getCellValue(cell))) {
							Coverage = getCellValue(cell);
						}
					}
					if(cellNum == 1) {
						if(!"".equals(getCellValue(cell))) {
							Category = getCellValue(cell);
						}
					}
					if(cellNum > 4) {
						data = new ArrayList<String>();
						data.add(benefit);
						data.add(Coverage);
						data.add(Category);
						data.add(planName.get(cellNum-5));
						if(cell.getCellType() == CellType.NUMERIC) {
							if (cell.getCellStyle().getDataFormatString().indexOf("%") > -1) {
								double num = cell.getNumericCellValue();
								String str = String.valueOf(num);
								int dg = str.length() - str.indexOf(".") - 1 - 2;
								num = num * 100;
								BigDecimal numberBigDecimal = new BigDecimal(num);
								data.add(numberBigDecimal + "%");
							} else {
								data.add(getCellValue(cell));
							}
						} else {
							data.add(getCellValue(cell));
						}
						dataList.add(data);
					}
				}
			}
		}
		System.out.println("[Benefit,  Coverage,  Category,  Plan Name,  Coverage Name]");
		System.out.println("===========================================================");
		for(List<String> datas : dataList) {
			for(int i = 0 ; i < datas.size() ; i++) {
				System.out.print(datas.get(i));
				if( i != (datas.size()-1)) {
					System.out.print(",");
				} else {
					System.out.println("");
				}
			}
			
		}
	}
	
	public static String getCellValue(Cell cell) {
		String cellValue = "";
		switch (cell.getCellType()) {
		case CellType.STRING:
			cellValue = cell.getStringCellValue();
			break;

		case CellType.FORMULA:
			cellValue = cell.getCellFormula();
			break;

		case CellType.NUMERIC:
			if (DateUtil.isCellDateFormatted(cell)) {
				cellValue = cell.getDateCellValue().toString();
			} else {
				cellValue = Double.toString(cell.getNumericCellValue());
			}
			break;

		case CellType.BLANK:
			cellValue = "";
			break;

		case CellType.BOOLEAN:
			cellValue = Boolean.toString(cell.getBooleanCellValue());
			break;

		}
		return cellValue;
	}
    
}

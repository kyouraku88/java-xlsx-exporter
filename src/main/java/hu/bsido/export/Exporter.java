package hu.bsido.export;

import java.io.File;
import java.io.FileOutputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.Collection;
import java.util.List;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Exporter {

	public static <T> void toXlsx(Collection<T> data, Class<T> klass, File file) {
		if (!ExportUtils.isExportable(klass)) {
			throw new IllegalArgumentException("the class must have annotation " + Exportable.class.getSimpleName());
		}

		List<Field> fields = ExportUtils.getExportableFields(klass);
		String sheetName = ExportUtils.getSheetName(klass);

		// Create a Workbook
		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
			CreationHelper createHelper = workbook.getCreationHelper();
			CellStyle dateCellStyle = workbook.createCellStyle();
			dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));

			// Create a Sheet
			Sheet sheet = workbook.createSheet(sheetName);

			// Create the header
			createHeaderRow(sheet, fields);

			int rowIdx = 1;
			for (T object : data) {
				Row row = sheet.createRow(rowIdx++);
				int cellIdx = 0;
				for (Field f : fields) {
					f.setAccessible(true);

					Cell cell = row.createCell(cellIdx++);
					Object fieldVal = f.get(object);

					if (f.getType().equals(Integer.class) || f.getType().equals(Integer.TYPE)) {
						cell.setCellValue(fieldVal == null ? 0d : Double.valueOf(fieldVal.toString()));

					} else if (f.getType().equals(String.class)) {
						cell.setCellValue((String) fieldVal);

					} else if (f.getType().equals(Boolean.class)) {
						cell.setCellValue(fieldVal == null ? false : (Boolean) fieldVal);

					} else if (f.getType().equals(Boolean.TYPE)) {
						cell.setCellValue((boolean) fieldVal);

					} else if (f.getType().isEnum()) {
						if (fieldVal != null) {
							cell.setCellValue(((Enum<?>) fieldVal).name());
						} else {
							cell.setCellValue((String) fieldVal);
						}
					}
				}
			}

			// Resize all columns to fit the content size
			for (int i = 0; i < fields.size(); i++) {
				sheet.autoSizeColumn(i);
			}

			// Write the output to a file
			FileOutputStream fileOut = new FileOutputStream(file);
			workbook.write(fileOut);
			fileOut.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static void createHeaderRow(Sheet sheet, List<Field> fields) {
		// Create a Row
		Row headerRow = sheet.createRow(0);

		// Create cells
		int fieldCount = 0;
		for (Field field : fields) {
			Cell cell = headerRow.createCell(fieldCount++);

			String header = null;
			for (Annotation a : field.getAnnotations()) {
				if (a instanceof Export) {
					Export ea = (Export) a;
					header = "".equals(ea.header()) ? field.getName() : ea.header();
				}
			}

			cell.setCellValue(header);
		}
	}
}

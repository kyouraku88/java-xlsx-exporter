package hu.bsido.export;

import java.io.File;
import java.io.FileOutputStream;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.Collection;
import java.util.HashSet;
import java.util.List;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CreationHelper;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import hu.bsido.common.Columns.StringColumn;
import hu.bsido.common.Export;
import hu.bsido.common.Exportable;
import hu.bsido.common.IEUtils;

public class Exporter {

	private static Set<Class<?>> primNumbers = new HashSet<>();

	static {
		primNumbers.add(Integer.TYPE);
		primNumbers.add(Double.TYPE);
		primNumbers.add(Float.TYPE);
		primNumbers.add(Short.TYPE);
		primNumbers.add(Long.TYPE);
	}

	private XSSFWorkbook workbook;
	private static CellStyle dateCellStyle;

	Exporter() {
		this.workbook = new XSSFWorkbook();
		CreationHelper createHelper = workbook.getCreationHelper();
		dateCellStyle = workbook.createCellStyle();
		dateCellStyle.setDataFormat(createHelper.createDataFormat().getFormat("dd-MM-yyyy"));
	}

	void createSheet(Class<?> klass, Collection<?> data, List<StringColumn> columns) {
		if (!IEUtils.isExportable(klass)) {
			throw new IllegalArgumentException("the class must have annotation " + Exportable.class.getSimpleName());
		}

		List<Field> fields = IEUtils.getExportableFields(klass);
		String sheetName = IEUtils.getSheetName(klass);

		// Create a Workbook
		try {
			// Create a Sheet
			Sheet sheet = workbook.createSheet(sheetName);

			// Create the header
			createHeaderRow(sheet, fields, columns);

			int rowIdx = 1;
			for (Object object : data) {
				Row row = sheet.createRow(rowIdx++);
				int cellIdx = 0;
				for (Field field : fields) {
					// make private fields accessible
					field.setAccessible(true);

					Cell cell = row.createCell(cellIdx++);
					Object fieldVal = field.get(object);
					setCellValue(cell, field, fieldVal);
				}

				if (columns != null) {
					for (StringColumn column : columns) {
						Cell cell = row.createCell(cellIdx);
						cell.setCellValue(column.getValue(object));
						cellIdx++;
					}
				}
			}

			// Resize all columns to fit the content size
			int columnCount = fields.size();
			if (columns != null) {
				columnCount += columns.size();
			}
			for (int i = 0; i < columnCount; i++) {
				sheet.autoSizeColumn(i);
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	void write(File file) {
		try {
			// Write the output to a file
			FileOutputStream fileOut = new FileOutputStream(file);
			workbook.write(fileOut);
			fileOut.close();
			workbook.close();
		} catch (Exception e) {
			e.printStackTrace();
		}
	}

	private static void setCellValue(Cell cell, Field field, Object fieldVal) {
		if (Number.class.isAssignableFrom(field.getType()) || primNumbers.contains(field.getType())) {
			cell.setCellValue(fieldVal == null ? 0d : Double.valueOf(fieldVal.toString()));

		} else if (field.getType().equals(String.class)) {
			cell.setCellValue((String) fieldVal);

		} else if (field.getType().equals(Boolean.class)) {
			cell.setCellValue(fieldVal == null ? false : (Boolean) fieldVal);

		} else if (field.getType().equals(Boolean.TYPE)) {
			cell.setCellValue((boolean) fieldVal);

		} else if (field.getType().equals(LocalDate.class)) {
			cell.setCellValue(IEUtils.convertToDate((LocalDate) fieldVal));
			cell.setCellStyle(dateCellStyle);

		} else if (field.getType().equals(LocalDateTime.class)) {
			cell.setCellValue(IEUtils.convertToDate((LocalDateTime) fieldVal));
			cell.setCellStyle(dateCellStyle);

		} else if (field.getType().isEnum()) {
			if (fieldVal != null) {
				cell.setCellValue(((Enum<?>) fieldVal).name());
			} else {
				cell.setCellValue((String) fieldVal);
			}

		} else if (Collection.class.isAssignableFrom(field.getType())) {
			cell.setCellValue(IEUtils.collectionToString((Collection<?>) fieldVal,
					field.getAnnotation(Export.class).separator()));
		}
	}

	private static void createHeaderRow(Sheet sheet, List<Field> fields, List<StringColumn> columns) {
		// Create a Row
		Row headerRow = sheet.createRow(0);

		// Create cells
		int cellIdx = 0;
		for (Field field : fields) {
			Cell cell = headerRow.createCell(cellIdx++);

			String header = field.getAnnotation(Export.class).header();
			if ("".equals(header)) {
				header = field.getName();
			}
			cell.setCellValue(header);
		}

		if (columns != null) {
			for (StringColumn column : columns) {
				Cell cell = headerRow.createCell(cellIdx);
				cell.setCellValue(column.getHeader());
				cellIdx++;
			}
		}
	}
}

package hu.bsido.inport;

import java.io.File;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Collection;
import java.util.Collections;
import java.util.HashMap;
import java.util.HashSet;
import java.util.Iterator;
import java.util.List;
import java.util.Map;
import java.util.Set;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

import hu.bsido.common.Export;
import hu.bsido.common.Exportable;
import hu.bsido.common.IEUtils;

public abstract class Importer {

	@SuppressWarnings({ "rawtypes", "unchecked" })
	public static List<?> fromXlsx(Class<?> klass, File from) {
		if (!IEUtils.isExportable(klass)) {
			throw new IllegalArgumentException("the class must have annotation " + Exportable.class.getSimpleName());
		}
		
		List<Field> fields = IEUtils.getExportableFields(klass);
		String sheetName = IEUtils.getSheetName(klass);
		
		List results = new ArrayList();
		try {
			Workbook workbook = WorkbookFactory.create(from);
			Sheet sheet = workbook.getSheet(sheetName);
			
			Iterator<Row> iter = sheet.iterator();
			Row headerRow = iter.next();

			Map<Integer, Field> columns = new HashMap<>();
			for (Field field : fields) {
				columns.put(getIndex(headerRow, field), field);
			}

			while (iter.hasNext()) {
				Row row = iter.next();
				Object o = klass.getConstructor().newInstance();
				
				for (Map.Entry<Integer, Field> entry : columns.entrySet()) {
					Cell cell = row.getCell(entry.getKey());
					Field field = entry.getValue();
					field.setAccessible(true);

					if (Number.class.isAssignableFrom(field.getType()) || IEUtils.primNumbers.contains(field.getType())) {
						Double doubleCellValue = cell.getNumericCellValue();

						if (Integer.class.equals(field.getType()) || Integer.TYPE.equals(field.getType())) {
							field.set(o, doubleCellValue.intValue());
							
						} else if (Double.class.equals(field.getType()) || Double.TYPE.equals(field.getType())) {
							field.set(o, doubleCellValue);
							
						} else if (Float.class.equals(field.getType()) || Float.TYPE.equals(field.getType())) {
							field.set(o, doubleCellValue.floatValue());
							
						} else if (Short.class.equals(field.getType()) || Short.TYPE.equals(field.getType())) {
							field.set(o, doubleCellValue.shortValue());
							
						} else if (Long.class.equals(field.getType()) || Long.TYPE.equals(field.getType())) {
							field.set(o, doubleCellValue.longValue());
						}
					} else if (String.class.equals(field.getType())) {
						field.set(o, cell.getStringCellValue());

					} else if (Boolean.class.equals(field.getType())
							|| Boolean.TYPE.equals(field.getType())) {
						field.set(o, cell.getBooleanCellValue());

					} else if (LocalDate.class.equals(field.getType())) {
						field.set(o, IEUtils.convertToLocalDate(cell.getDateCellValue()));

					} else if (LocalDateTime.class.equals(field.getType())) {
						field.set(o, IEUtils.convertToLocalDateTime(cell.getDateCellValue()));

					} else if (field.getType().isEnum()) {
						String enumVal = cell.getStringCellValue();
						if (enumVal != null) {
							field.set(o, Enum.valueOf((Class<? extends Enum>) Class.forName(field.getType().getName()), enumVal));
						}

					} else if (Collection.class.isAssignableFrom(field.getType())) {
						String separator = field.getAnnotation(Export.class).separator();
						Class<?> type = field.getAnnotation(Export.class).collectionType();
						List<String> values = Arrays.asList(cell.getStringCellValue().split(separator));

						if (Set.class.equals(field.getType())) {
							Set set = Collections.checkedSet(new HashSet(), type);
							if (String.class.equals(type)) {
								set.addAll(values);
							} else if (Integer.class.equals(type)) {
								for (String value : values) {
									set.add(Integer.valueOf(value));
								}
							}
							field.set(o, set);

						} else if (List.class.equals(field.getType())) {
							List list = Collections.checkedList(new ArrayList(), type);
							if (String.class.equals(type)) {
								list.addAll(values);
							} else if (Integer.class.equals(type)) {
								for (String value : values) {
									list.add(Integer.valueOf(value));
								}
							}
							field.set(o, list);
						}
					}
				}
				results.add(o);
			}
			
		} catch (Exception e) {
			e.printStackTrace();
		}
		
		return results;
	}

	private static Integer getIndex(Row headerRow, Field field) {
		String header = field.getAnnotation(Export.class).header();
		if ("".equals(header)) {
			header = field.getName();
		}

		for (int cellIdx = 0; cellIdx < headerRow.getLastCellNum(); cellIdx++) {
			String cellValue = headerRow.getCell(cellIdx).getStringCellValue();
			if (header.equals(cellValue)) {
				return cellIdx;
			}
		}
		throw new IllegalStateException("the column was not found in the file!");
	}

}

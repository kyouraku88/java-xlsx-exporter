package hu.bsido.common;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.ZoneId;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.List;

public abstract class IEUtils {

	public static LocalDate convertToLocalDate(Date dateToConvert) {
		if (dateToConvert == null) {
			return null;
		}
		return new java.sql.Date(dateToConvert.getTime()).toLocalDate();
	}

	public static LocalDateTime convertToLocalDateTime(Date dateToConvert) {
		if (dateToConvert == null) {
			return null;
		}
		return new java.sql.Timestamp(dateToConvert.getTime()).toLocalDateTime();
	}

	public static Date convertToDate(LocalDate dateToConvert) {
		if (dateToConvert == null) {
			return null;
		}
		return java.util.Date.from(dateToConvert.atStartOfDay().atZone(ZoneId.systemDefault()).toInstant());
	}

	public static Date convertToDate(LocalDateTime dateToConvert) {
		if (dateToConvert == null) {
			return null;
		}
		return java.sql.Timestamp.valueOf(dateToConvert);
	}

	public static boolean isExportable(Class<?> klass) {
		boolean exportable = false;
		for (Annotation a : klass.getAnnotations()) {
			if (a instanceof Exportable) {
				exportable = true;
			}
		}
		return exportable;
	}

	public static List<Field> getExportableFields(Class<?> klass) {
		List<Field> result = new ArrayList<>();

		for (Field field : klass.getDeclaredFields()) {
			for (Annotation a : field.getDeclaredAnnotations()) {
				if (a instanceof Export) {
					result.add(field);
				}
			}
		}

		return result;
	}

	public static String getSheetName(Class<?> klass) {
		String result = null;
		for (Annotation a : klass.getAnnotations()) {
			if (a instanceof Exportable) {
				Exportable ea = (Exportable) a;
				result = "".equals(ea.sheet()) ? klass.getSimpleName() : ea.sheet();
			}
		}
		// TODO result cannot be null

		return result;
	}

	public static String collectionToString(Collection<?> collection, String separator) {
		if (collection == null || collection.isEmpty()) {
			return "";
		}

		StringBuilder sb = new StringBuilder();
		for (Object item : collection) {
			sb.append(item.toString());
			sb.append(separator);
		}
		// remove the last separator
		return sb.substring(0, sb.length() - separator.length());
	}
}

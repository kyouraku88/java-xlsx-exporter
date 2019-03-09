package hu.bsido.export;

import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

public abstract class ExportUtils {

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
}

package hu.bsido.export;

import java.io.File;
import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;

import hu.bsido.common.Columns.StringColumn;

public class ExportBuilder {

	private final Exporter exporter;
	private final Map<Class<?>, Collection<?>> sheets;
	private final Map<Class<?>, List<StringColumn>> columns;

	private ExportBuilder() {
		exporter = new Exporter();
		sheets = new LinkedHashMap<>();
		columns = new LinkedHashMap<>();
	}

	public static ExportBuilder create() {
		return new ExportBuilder();
	}

	public ExportBuilder addSheet(Class<?> klass, Collection<?> list) {
		sheets.put(klass, list);
		return this;
	}

	public void export(File file) {
		for (Map.Entry<Class<?>, Collection<?>> entry : sheets.entrySet()) {
			exporter.createSheet(entry.getKey(), entry.getValue(), columns.get(entry.getKey()));
		}
		exporter.write(file);
	}

	public ExportBuilder addColumn(Class<?> klass, StringColumn column) {
		if (columns.get(klass) == null) {
			columns.put(klass, new ArrayList<StringColumn>());
		}
		columns.get(klass).add(column);
		return this;
	}
}

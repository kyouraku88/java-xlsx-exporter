package hu.bsido.export;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

public class ExportTest {
	
	public enum TestEnum {
		VAL1, VAL2
	}
	
	@Exportable(sheet = "Exportable")
	private static class ExportableModel {
		@Export
		private int exportValue;
		@SuppressWarnings("unused")
		private int dontExport;
		@Export
		private TestEnum enumValue;
		@Export
		private String stringExport;
		
		public ExportableModel(int exportValue, int dontExport) {
			this.exportValue = exportValue;
			this.dontExport = dontExport;
			enumValue = TestEnum.VAL1;
			stringExport = "str";
		}
	}
	
	@Exportable
	private static class ExportableModel2 {
	}
	
	private static class NotExporttable {}

	@Test
	public void testField() throws NoSuchFieldException, SecurityException {
		Class<ExportableModel> cem = ExportableModel.class;
		List<Field> fields = ExportUtils.getExportableFields(cem);
		Field exportedField = cem.getDeclaredField("exportValue");
		assertEquals(exportedField, fields.get(0));
	}
	
	@Test
	public void testSheetFromValue() {
		Class<ExportableModel> cem = ExportableModel.class;
		String sheetName = ExportUtils.getSheetName(cem);
		assertEquals("Exportable", sheetName);
	}
	
	@Test
	public void testSheetFromClass() {
		Class<ExportableModel2> cem = ExportableModel2.class;
		String sheetName = ExportUtils.getSheetName(cem);
		assertEquals("ExportableModel2", sheetName);
	}
	
	@Test
	public void testExportToXlsx() {
		List<ExportableModel> list = new ArrayList<>();
		list.add(new ExportableModel(5, 10));
		Exporter.toXlsx(list, ExportableModel.class, new File("test.xlsx"));
	}
	
	@Test
	public void testExportable() {
		Class<ExportableModel> cex = ExportableModel.class;
		assertEquals(true, ExportUtils.isExportable(cex));
		
		Class<NotExporttable> cnotex = NotExporttable.class;
		assertEquals(false, ExportUtils.isExportable(cnotex));
	}
}

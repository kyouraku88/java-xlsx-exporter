package hu.bsido.export;

import static org.junit.Assert.assertEquals;

import java.io.File;
import java.lang.reflect.Field;
import java.util.ArrayList;
import java.util.List;

import org.junit.Test;

import hu.bsido.common.Export;
import hu.bsido.common.Exportable;
import hu.bsido.common.IEUtils;

public class ExportTest {
	
	public enum TestEnum {
		VAL1, VAL2
	}
	
	@Exportable(sheet = "Exportable")
	private static class ExportableModel {
		@Export(header = "intExport")
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
		List<Field> fields = IEUtils.getExportableFields(cem);
		Field exportedField = cem.getDeclaredField("exportValue");
		assertEquals(exportedField, fields.get(0));
	}
	
	@Test
	public void testSheetFromValue() {
		Class<ExportableModel> cem = ExportableModel.class;
		String sheetName = IEUtils.getSheetName(cem);
		assertEquals("Exportable", sheetName);
	}
	
	@Test
	public void testSheetFromClass() {
		Class<ExportableModel2> cem = ExportableModel2.class;
		String sheetName = IEUtils.getSheetName(cem);
		assertEquals("ExportableModel2", sheetName);
	}

	@Test
	public void testExportToXlsx() {
		List<ExportableModel> list = new ArrayList<>();
		list.add(new ExportableModel(5, 10));
		ExportBuilder.create().addSheet(ExportableModel.class, list).export(new File("test.xlsx"));
	}

	@Test
	public void testExportable() {
		Class<ExportableModel> cex = ExportableModel.class;
		assertEquals(true, IEUtils.isExportable(cex));
		
		Class<NotExporttable> cnotex = NotExporttable.class;
		assertEquals(false, IEUtils.isExportable(cnotex));
	}
}

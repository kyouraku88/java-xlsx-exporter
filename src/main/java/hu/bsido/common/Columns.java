package hu.bsido.common;

public abstract class Columns {

	private static abstract class Column<T> {
		private final String header;
		public Column(String header) {
			this.header = header;
		}
		public String getHeader() {
			return header;
		}
		public abstract T getValue(Object object);
	}

	public static abstract class StringColumn extends Column<String> {
		public StringColumn(String header) {
			super(header);
		}
	}
}

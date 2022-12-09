package statsProcessor;

class ExcelExtractorException extends Exception {
	static final long serialVersionUID = 43L;
	
	public ExcelExtractorException(String m, int rowAt) {
		super("ROWAT<" + rowAt + ">-" + m );
	}
}
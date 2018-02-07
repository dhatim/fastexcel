package org.dhatim.fastexcel.reader;

@SuppressWarnings("serial")
public class ExcelReaderException extends RuntimeException {

    public ExcelReaderException(String message) {
        super(message);
    }

    public ExcelReaderException(String message, Throwable cause) {
        super(message, cause);
    }

    public ExcelReaderException(Throwable cause) {
        super(cause);
    }

}

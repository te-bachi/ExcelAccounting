package ch.fablabwinti.accounting.cell;

/**
 *
 */
public class CustomCellException extends Exception {
    public CustomCellException() {
        super();
    }

    public CustomCellException(String message) {
        super(message);
    }

    public CustomCellException(String message, Throwable cause) {
        super(message, cause);
    }

    public CustomCellException(Throwable cause) {
        super(cause);
    }

    protected CustomCellException(String message, Throwable cause, boolean enableSuppression, boolean writableStackTrace) {
        super(message, cause, enableSuppression, writableStackTrace);
    }
}

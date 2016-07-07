package ch.fablabwinti.accounting;

import java.math.BigDecimal;
import java.util.Date;

/**
 *
 */
public interface JournalFilter {
    public boolean filter(int nr, Date date, int debit, int credit, BigDecimal amount, String text);
}

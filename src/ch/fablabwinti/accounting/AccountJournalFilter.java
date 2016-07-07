package ch.fablabwinti.accounting;

import java.math.BigDecimal;
import java.util.Date;

/**
 *
 */
public class AccountJournalFilter implements JournalFilter {

    private int account;

    public AccountJournalFilter(int account) {
        this.account = account;
    }

    @Override
    public boolean filter(int nr, Date date, int debit, int credit, BigDecimal amount, String text) {
        if (account == debit || account == credit) {
            return true;
        }

        return false;
    }
}

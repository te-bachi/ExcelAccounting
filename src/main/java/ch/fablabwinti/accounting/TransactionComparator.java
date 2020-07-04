package ch.fablabwinti.accounting;

import java.util.Comparator;

/**
 *
 */
public class TransactionComparator implements Comparator<Transaction> {
    @Override
    public int compare(Transaction t1, Transaction t2) {
        if (t1.getDate().equals(t2.getDate())) {
            return Integer.valueOf(t1.getNr()).compareTo(Integer.valueOf(t2.getNr()));
        }
        return t1.getDate().compareTo(t2.getDate());
    }
}

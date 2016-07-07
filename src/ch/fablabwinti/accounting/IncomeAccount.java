package ch.fablabwinti.accounting;

/**
 *
 */
public class IncomeAccount extends Account {
    public IncomeAccount(Account parent, int number, String name) {
        super(parent, number, name);
    }

    public IncomeAccount(int number, String name) {
        super(number, name);
    }
}

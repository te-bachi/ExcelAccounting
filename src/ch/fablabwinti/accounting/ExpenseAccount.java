package ch.fablabwinti.accounting;

/**
 *
 */
public class ExpenseAccount extends Account {
    public ExpenseAccount(Account parent, int number, String name) {
        super(parent, number, name);
    }

    public ExpenseAccount(int number, String name) {
        super(number, name);
    }
}

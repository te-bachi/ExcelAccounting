package ch.fablabwinti.accounting;

/**
 *
 */
public class TitleAccount extends Account {
    public TitleAccount(Account parent, int number, String name) {
        super(parent, number, name);
    }

    public TitleAccount(int number, String name) {
        super(number, name);
    }
}

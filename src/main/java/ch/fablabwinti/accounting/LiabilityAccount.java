package ch.fablabwinti.accounting;

/**
 *
 */
public class LiabilityAccount extends Account {
    public LiabilityAccount(Account parent, int number, String name) {
        super(parent, number, name);
    }

    public LiabilityAccount(int number, String name) {
        super(number, name);
    }
}

package ch.fablabwinti.accounting;

/**
 *
 */
public class AssetAccount extends Account {
    public AssetAccount(Account parent, int number, String name) {
        super(parent, number, name);
    }

    public AssetAccount(int number, String name) {
        super(number, name);
    }
}

package ch.fablabwinti.accounting;

public class AccountNotFoundException extends Exception {
    private AccountNotFoundException() {
        super();
    }

    private AccountNotFoundException(String message) {
        super(message);
    }

    public AccountNotFoundException(int accountNr) {
        super("Account number \"" + accountNr + "\" not found");
    }
}

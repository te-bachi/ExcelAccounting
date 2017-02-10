package ch.fablabwinti.accounting;

import java.util.ArrayList;
import java.util.List;

/**
 *
 */
public class AccountList {
    private List<Account> accountList;

    public AccountList() {
        this.accountList = new ArrayList();
    }

    public AccountList(List<Account> accountList) {
        this.accountList = accountList;
    }

    public void add(Account account) {
        accountList.add(account);
    }

    public Account get(int index) {
        return accountList.get(index);
    }

    public int size() {
        return accountList.size();
    }

    public boolean isEmpty() {
        return accountList.isEmpty();
    }

    public Account removeNext() {
        return accountList.remove(0);
    }

    public Account find(int accountNr) {
        Account match = null;
        for (Account account : accountList) {
            if (account.getNumber() == accountNr) {
                match = account;
                break;
            }
        }
        return match;
    }

    public Account remove(int accountNr) {
        Account match = null;
        Account account;
        int     i;

        for (i = 0; i < accountList.size(); i++) {
            account = accountList.get(i);
            if (account.getNumber() == accountNr) {
                match = accountList.remove(i);
                break;
            }
        }
        return match;
    }

}

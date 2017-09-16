package ch.fablabwinti.accounting;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;
import java.util.Spliterator;
import java.util.function.Consumer;

/**
 *
 */
public class AccountList implements Iterable<Account> {
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

    /**
     * Find account number only in the first layer, not in the deeper layers
     *
     * @param accountNr
     * @return
     * @throws AccountNotFoundException
     */
    public Account find(int accountNr) throws AccountNotFoundException {
        Account match = null;
        for (Account account : accountList) {
            if (account.getNumber() == accountNr) {
                match = account;
                break;
            }
        }

        if (match == null) {
            throw new AccountNotFoundException(accountNr);
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

    @Override
    public Iterator<Account> iterator() {
        return accountList.iterator();
    }
}

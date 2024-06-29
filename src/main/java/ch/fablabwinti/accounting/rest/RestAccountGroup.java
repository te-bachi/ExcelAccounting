package ch.fablabwinti.accounting.rest;

import java.util.ArrayList;
import java.util.List;

public class RestAccountGroup extends RestObject {

    private List<RestAccount> accountList;

    public RestAccountGroup() {
        this.accountList = new ArrayList<>();
    }

    public List<RestAccount> getAccountList() {
        return accountList;
    }

    public void setAccountList(List<RestAccount> accountList) {
        this.accountList = accountList;
    }
}

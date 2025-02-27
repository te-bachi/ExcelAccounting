package ch.fablabwinti.accounting.rest;

import com.fasterxml.jackson.annotation.JsonIgnoreProperties;

import java.util.ArrayList;
import java.util.List;

@JsonIgnoreProperties(ignoreUnknown = true)
public class RestPeriod extends RestObject {
    private List<RestEntrygroup> entryGroupList;
    private List<RestAccountGroup> accountGroupList;
    private List<RestCostcenter> costcenterList;

    public RestPeriod() {
        this.entryGroupList = new ArrayList<>();
        this.accountGroupList = new ArrayList<>();
        this.costcenterList = new ArrayList<>();
    }

    public List<RestEntrygroup> getEntryGroupList() {
        return entryGroupList;
    }

    public void setEntryGroupList(List<RestEntrygroup> entryGroupList) {
        this.entryGroupList = entryGroupList;
    }

    public List<RestAccountGroup> getAccountGroupList() {
        return accountGroupList;
    }

    public void setAccountGroupList(List<RestAccountGroup> accountGroupList) {
        this.accountGroupList = accountGroupList;
    }

    public List<RestCostcenter> getCostcenterList() {
        return costcenterList;
    }

    public void setCostcenterList(List<RestCostcenter> costcenterList) {
        this.costcenterList = costcenterList;
    }
}

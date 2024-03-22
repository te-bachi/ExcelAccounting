package ch.fablabwinti.accounting.rest;

import java.util.Map;

public class RestObject {
    private String type;
    private boolean readonly;

    private Map<String, String> properties;

    private Map<String, int[]> children;

    private Map<String, int[]> links;

    private int[] parents;

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public boolean isReadonly() {
        return readonly;
    }

    public void setReadonly(boolean readonly) {
        this.readonly = readonly;
    }

    public Map<String, String> getProperties() {
        return properties;
    }

    public void setProperties(Map<String, String> properties) {
        this.properties = properties;
    }

    public Map<String, int[]> getChildren() {
        return children;
    }

    public void setChildren(Map<String, int[]> children) {
        this.children = children;
    }

    public Map<String, int[]> getLinks() {
        return links;
    }

    public void setLinks(Map<String, int[]> links) {
        this.links = links;
    }

    public int[] getParents() {
        return parents;
    }

    public void setParents(int[] parents) {
        this.parents = parents;
    }
}

package ch.fablabwinti.accounting.rest;

public class RestObjectList {
    private int[] objects;

    private String error;

    public int[] getObjects() {
        return objects;
    }

    public void setObjects(int[] objects) {
        this.objects = objects;
    }

    public String getError() {
        return error;
    }

    public void setError(String error) {
        this.error = error;
    }
}

package com.kkllor.exceldata.entity;

public class ResultBean {

    private String originData;
    private String name;
    private String type;
    private String range;
    private long index;

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public String getType() {
        return type;
    }

    public void setType(String type) {
        this.type = type;
    }

    public String getRange() {
        return range;
    }

    public void setRange(String range) {
        this.range = range;
    }

    public String getOriginData() {
        return originData;
    }

    public void setOriginData(String originData) {
        this.originData = originData;
    }

    public long getIndex() {
        return index;
    }

    public void setIndex(long index) {
        this.index = index;
    }

    @Override
    public String toString() {
        return "ResultBean{" +
                "originData='" + originData + '\'' +
                ", name='" + name + '\'' +
                ", type='" + type + '\'' +
                ", range='" + range + '\'' +
                ", index=" + index +
                '}';
    }

    public boolean isValidate() {
        if (originData == null || name == null || range == null || type == null) {
            return false;
        }
        return true;
    }
}

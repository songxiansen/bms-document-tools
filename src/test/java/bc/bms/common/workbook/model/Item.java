package bc.bms.common.workbook.model;

import java.io.Serializable;

public class Item implements Editable, Serializable {

    private Integer id;

    private String name;

    private boolean editable;

    public Item() {

    }

    public Item(Integer id, String name) {
        this.id = id;
        this.name = name;
    }

    public Item(Integer id, String name, boolean editable) {
        this.id = id;
        this.name = name;
        this.editable = editable;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public boolean isEditable() {
        return editable;
    }

    public void setEditable(boolean editable) {
        this.editable = editable;
    }

    @Override
    public String toString() {
        return "Item{" +
                "id=" + id +
                ", name='" + name + '\'' +
                ", editable=" + editable +
                '}';
    }
}

package bc.bms.common.workbook.model;

public enum ProjectStatus {

    Draft(1, "正在交易中"),
    Approved(2, "已成功交易"),
    Adjusting(3, "交易退款中"),
    Rejected(4, "交易被拒绝"),
    freeze(5,"交易被冻结");
    
    private int id;

    private String name;

    private ProjectStatus(int id, String name) {
        this.id = id;
        this.name = name;
    }

    public int getId() {
        return id;
    }

    public String getCode() {
        return this.toString();
    }

    public String getName() {
        return name;
    }

    public static ProjectStatus getProjectStatus(int id) {
        for (ProjectStatus projectStatus : ProjectStatus.values()) {
            if (projectStatus.getId() == (id % 5)) {
                return projectStatus;
            }
        }

        return Draft;
    }

}

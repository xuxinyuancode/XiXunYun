
import com.alibaba.excel.annotation.ExcelProperty;
import lombok.Data;


@Data
public class Model {
    @ExcelProperty(value = "填报时间")
    private String writeTime;

    /**
     * 实习工作具体情况及实习任务完成情况
     * "今日工作总结","本周工作总结","本月总结"
     */
    @ExcelProperty({"今日工作总结"})
    private String sumText;


    /**
     * 主要收获及工作成绩
     */
    @ExcelProperty(value = "主要收获及工作成绩")
    private String gainText;

    /**
     * 工作中的问题及需要老师的指导帮助
     */
    @ExcelProperty(value = "工作中的问题及需要老师的指导帮助")
    private String uselessText;



}

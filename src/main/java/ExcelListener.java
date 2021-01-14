import com.alibaba.excel.context.AnalysisContext;
import com.alibaba.excel.event.AnalysisEventListener;

public class ExcelListener extends AnalysisEventListener<Model> {
    public void invoke(Model model, AnalysisContext analysisContext) {
        if (model.getWriteTime() == null) {
            return;
        }
        Model m = model;
        System.out.println("****时间：" + m.getWriteTime());
        if (model.getSumText() != null) {
            m.setSumText(model.getSumText().replaceAll("\\pS|\\pC", ""));
        }
        System.out.println("****总结：" + m.getSumText());
        if (model.getGainText() != null) {
            m.setGainText(model.getGainText().replaceAll("\\pS|\\pC", ""));
        }
        System.out.println("****收获：" + m.getGainText());

        if (model.getUselessText() != null) {
            m.setUselessText(model.getUselessText().replaceAll("\\pS|\\pC", ""));
        }
        System.out.println("****问题：" + m.getUselessText());

        System.out.println("-----------------------------------");
        System.out.println();

        GlobalVariables.DATA.add(m);
    }


    public void doAfterAllAnalysed(AnalysisContext analysisContext) {
        System.out.println("数据读取完成");
    }

}
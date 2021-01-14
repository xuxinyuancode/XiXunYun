import com.alibaba.excel.EasyExcel;
import com.alibaba.excel.ExcelReader;
import com.alibaba.excel.annotation.ExcelProperty;
import org.jsoup.Connection;
import org.jsoup.Jsoup;
import org.jsoup.nodes.Document;

import java.io.File;
import java.io.IOException;


import java.lang.annotation.Annotation;
import java.lang.reflect.Field;
import java.lang.reflect.InvocationHandler;
import java.lang.reflect.Proxy;
import java.net.URLEncoder;
import java.util.Arrays;
import java.util.Map;
import java.util.Scanner;

public class Launch {
    public static void main(String[] args) {
        Scanner scanner = new Scanner(System.in);
        System.out.println("请输入Token:");
        GlobalVariables.TOKEN = scanner.nextLine();
        System.out.println("您的Token为：" + GlobalVariables.TOKEN);
        System.out.println("请输入日志路径:");
        GlobalVariables.PATH = scanner.nextLine();
        System.out.println("您的日志路径为：" + GlobalVariables.PATH);

        File file = new File(GlobalVariables.PATH);
        if (!file.exists()) {
            System.out.println("日志不存在，请重试");
            return;
        }

        System.out.println("请输您要上传的类型: 1为日报 2为周报 3为月报");
        GlobalVariables.TYPE = scanner.nextLine();
        System.out.println("您要上传的类型为：" + GlobalVariables.TYPE);


        GlobalVariables.HEADER.put("Connection", "keep-alive");
        GlobalVariables.HEADER.put("Accept", "application/json, text/javascript, */*; q=0.01'");
        GlobalVariables.HEADER.put("User-Agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/87.0.4280.141 Safari/537.36 Edg/87.0.664.75");
        GlobalVariables.HEADER.put("Content-Type", "application/x-www-form-urlencoded; charset=UTF-8");
        GlobalVariables.HEADER.put("Origin", "https://www.xixunyun.com");
        GlobalVariables.HEADER.put("Sec-Fetch-Site", "same-site");
        GlobalVariables.HEADER.put("Sec-Fetch-Mode", "cors");
        GlobalVariables.HEADER.put("Sec-Fetch-Dest", "empty");
        GlobalVariables.HEADER.put("Referer", "https://www.xixunyun.com");
        GlobalVariables.HEADER.put("Accept-Language", "zh-CN,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6");
        ExcelReader excelReader = null;
        GlobalVariables.DATA.clear();


        try {
            String field = "今日工作总结";
            if (GlobalVariables.TYPE.contains("1")) {
                field = "今日工作总结";
            } else if (GlobalVariables.TYPE.contains("2")) {
                field = "本周工作总结";
            } else if (GlobalVariables.TYPE.contains("3")) {
                field = "本月总结";
            }
            System.out.println(field);
            Class<?> modelClass = changeAnnotationValue(Model.class, ExcelProperty.class, "value", new String[]{field});
            try {
                Field[] fields = modelClass.getDeclaredFields();
                for (Field field1 : fields) {
                    ExcelProperty annotation = field1.getAnnotation(ExcelProperty.class);
                    System.out.println(Arrays.toString(annotation.value()));
                }
            } catch (Exception e) {
                e.printStackTrace();
            }
            EasyExcel.read(GlobalVariables.PATH, modelClass, new ExcelListener()).sheet().doRead();
        } finally {
            if (excelReader != null) {
                // 这里千万别忘记关闭，读的时候会创建临时文件，到时磁盘会崩的
                excelReader.finish();
            }
        }
        for (int i = 0; i < GlobalVariables.DATA.size(); i++) {
            try {
                Thread.sleep(1000);
            } catch (InterruptedException e) {
                e.printStackTrace();
            }
            String s = sendPost(GlobalVariables.TOKEN, GlobalVariables.DATA.get(i));
            if (s.contains("20000")) {
                System.out.println(GlobalVariables.DATA.get(i).getWriteTime() + ": 成功");
            } else {
                System.out.println(GlobalVariables.DATA.get(i).getWriteTime() + ": 失败  " + s );
            }
        }


    }


    /**
     * 向指定 URL 发送POST方法的请求
     *
     * @param token 发送请求的 token
     * @param model 请求参数
     * @return 所代表远程资源的响应结果
     */
    public static String sendPost(String token, Model model) {
        String url = "https://api.xixunyun.com/Reports/StudentOperator?token=" + token;
        try {
            Connection connection = Jsoup.connect(url)
                    .headers(GlobalVariables.HEADER).ignoreContentType(true);
            if (GlobalVariables.TYPE.equals("1"))
                connection.data("business_type", "day");
            else if (GlobalVariables.TYPE.equals("2"))
                connection.data("business_type", "week");
            else if (GlobalVariables.TYPE.equals("3"))
                connection.data("business_type", "month");
            else return "";
            String time = model.getWriteTime().substring(0, 10).replace("年", "/").replace("月", "/");
            connection.data("start_date", time);
            connection.data("end_date", time);
            connection.data("content", "[{\"title\":\"实习工作具体情况及实习任务完成情况\",\"content\":\"" + model.getSumText() + "\",\"require\":\"1\",\"sort\":1},{\"title\":\"主要收获及工作成绩\",\"content\":\"" + model.getGainText() + "\",\"require\":\"0\",\"sort\":2},{\"title\":\"工作中的问题及需要老师的指导帮助\",\"content\":\"" + model.getUselessText() + "\",\"require\":\"0\",\"sort\":3}]");
            connection.data("attachment", "");
            Document doc = connection.post();
            return doc.toString();
        } catch (IOException e) {
            e.printStackTrace();
        }
        return "";

    }

    /**
     * 变更注解的属性值
     *
     * @param clazz     注解所在的实体类
     * @param tClass    注解类
     * @param filedName 要修改的注解属性名
     * @param value     要设置的属性值
     */
    public static <A extends Annotation> Class<?> changeAnnotationValue(Class<?> clazz, Class<A> tClass, String filedName, Object value) {
        try {
            Field[] fields = clazz.getDeclaredFields();
            for (Field field : fields) {
                if (field.getName().equals("sumText")) {
                    A annotation = field.getAnnotation(tClass);
                    setAnnotationValue(annotation, filedName, value);
                }
            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return clazz;
    }

    /**
     * 设置注解中的字段值
     *
     * @param annotation 要修改的注解实例
     * @param fieldName  要修改的注解属性名
     * @param value      要设置的属性值
     */
    public static void setAnnotationValue(Annotation annotation, String fieldName, Object value) throws NoSuchFieldException, IllegalAccessException {
        InvocationHandler handler = Proxy.getInvocationHandler(annotation);
        Field field = handler.getClass().getDeclaredField("memberValues");
        field.setAccessible(true);
        Map memberValues = (Map) field.get(handler);
        memberValues.put(fieldName, value);
    }


}

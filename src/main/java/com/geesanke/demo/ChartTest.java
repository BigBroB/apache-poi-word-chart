package com.geesanke.demo;

import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.config.Configure;
import com.geesanke.demo.data.LineChartRenderData;
import com.geesanke.demo.policy.InsertLineChartRefRenderPolicy;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/**
 * @ClassName ChartTest
 * @Description 测试
 * @Author yeehaw
 * @Date 2020/4/20 14:25
 * @Version 1.0.0
 */
public class ChartTest {

    public static String templatePath = "yours/test.docx";
    public static String after = "yours/test_after.doc";



    public static void main(String[] args) throws Exception {
        LineChartRenderData data = new LineChartRenderData();
        List<String> x = new ArrayList<String>();
        x.add("A");
        x.add("B");
        data.setPointX(x);
        data.setTitle("testABC");
        List<LineChartRenderData.LineData> lines = new ArrayList<LineChartRenderData.LineData>();
        LineChartRenderData.LineData l1 = new LineChartRenderData.LineData();
        l1.setTitle("test1");
        Map<String, BigDecimal> l1points = new HashMap<String, BigDecimal>();
        l1points.put("A", new BigDecimal(1.2));
        l1points.put("B", new BigDecimal(5));
        l1.setPoints(l1points);
        lines.add(l1);

        LineChartRenderData.LineData l2 = new LineChartRenderData.LineData();
        l2.setTitle("test2");
        Map<String, BigDecimal> l2points = new HashMap<String, BigDecimal>();
        l2points.put("A", new BigDecimal(2.0));
        l2points.put("B", new BigDecimal(3.5));
        l2.setPoints(l2points);
        lines.add(l2);
        data.setLines(lines);


        Configure configure = Configure.newBuilder()
                .referencePolicy(
                        new InsertLineChartRefRenderPolicy("列1", data))
                .build();

        Map<String, Object> params = new HashMap<String, Object>();
        params.put("text1", "测试数据1 \n 测试空行");
        XWPFTemplate template = XWPFTemplate
                .compile(templatePath, configure)
                .render(params);
        template.writeToFile(after);
    }
}

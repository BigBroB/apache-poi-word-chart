package com.geesanke.demo.policy;

import com.deepoove.poi.XWPFTemplate;
import com.geesanke.demo.data.LineChartRenderData;
import com.geesanke.demo.util.LineChartRenderUtil;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.apache.poi.xwpf.usermodel.XWPFChart;

import java.util.List;

/**
 * @ClassName ReplaceOptionalChartRefRenderPolicy
 * @Description 根据可选文字替换图表内容
 * @Author yeehaw
 * @Date 2020/4/20 13:47
 * @Version 1.0.0
 */
@Getter
@Setter
@NoArgsConstructor
public class InsertLineChartRefRenderPolicy extends OptionalChartRefRenderPolicy {

    private LineChartRenderData data;
    private String signal;

    public InsertLineChartRefRenderPolicy(String signal, LineChartRenderData lineChartRenderData) {
        this.data = lineChartRenderData;
        this.signal = signal;
    }

    @Override
    public String optionalText() {
        return this.signal;
    }

    @Override
    public void doRender(List<XWPFChart> charts, XWPFTemplate template) throws Exception {
        for (XWPFChart chart : charts) {
            LineChartRenderUtil.renderExcel(chart, data);
            LineChartRenderUtil.renderChart(chart, data);
        }
    }


}

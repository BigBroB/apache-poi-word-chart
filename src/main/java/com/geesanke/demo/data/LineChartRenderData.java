package com.geesanke.demo.data;

import lombok.Data;

import java.math.BigDecimal;
import java.util.List;
import java.util.Map;

/**
 * @ClassName LineChartRenderData
 * @Description 折线图数据
 * @Author yeehaw
 * @Date 2020/4/21 14:09
 * @Version 1.0.0
 */
@Data
public class LineChartRenderData {
    /**
     * 图名
     */
    private String title;
    /**
     * 折线
     */
    private List<LineData> lines;
    /**
     * 横坐标
     */
    private List<String> pointX;


    @Data
    public static class LineData {
        /**
         * excel title 折线图 标题
         */
        private String title;
        /**
         * 值
         */
        private Map<String, BigDecimal> points;

    }
}

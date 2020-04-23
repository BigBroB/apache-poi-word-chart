package com.geesanke.demo.policy;

import cn.hutool.core.util.StrUtil;
import com.deepoove.poi.XWPFTemplate;
import com.deepoove.poi.policy.ref.OptionalText;
import com.deepoove.poi.policy.ref.ReferenceRenderPolicy;
import com.deepoove.poi.xwpf.NiceXWPFDocument;
import com.google.common.collect.Lists;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFChart;

import java.util.List;

/**
 * @ClassName OptionalChartRefRenderPolicy
 * @Description 可选文字匹配 XWPFChart
 * @Author yeehaw
 * @Date 2020/4/20 13:16
 * @Version 1.0.0
 */
public abstract class OptionalChartRefRenderPolicy extends ReferenceRenderPolicy<List<XWPFChart>> implements OptionalText {

    @Override
    protected List<XWPFChart> locate(XWPFTemplate template) {
        logger.info("Try locate the XWPFChart object which mathing optional text [{}]...", optionalText());
        try {
            List<XWPFChart> findCharts = Lists.newArrayList();
            NiceXWPFDocument document = template.getXWPFDocument();
            // 获取所有图表
            List<XWPFChart> charts = document.getCharts();
            for (XWPFChart chart : charts) {
                XSSFWorkbook workbook = chart.getWorkbook();
                XSSFSheet sheet1 = workbook.getSheetAt(0);
                XSSFCell cell = sheet1.getRow(0).getCell(0);
                String firstCellValue = cell.getStringCellValue();
                if (StrUtil.equals(firstCellValue, optionalText())) {
                    findCharts.add(chart);
                }
            }
            return findCharts;
        } catch (Exception e) {

        }
        return null;
    }


}

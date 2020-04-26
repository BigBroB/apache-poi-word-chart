package com.geesanke.demo.util;

import cn.hutool.core.util.ReflectUtil;
import com.geesanke.demo.data.LineChartRenderData;
import org.apache.poi.ooxml.POIXMLDocumentPart;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xddf.usermodel.chart.XDDFChartData;
import org.apache.poi.xddf.usermodel.chart.XDDFLineChartData;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.apache.poi.xwpf.usermodel.XWPFChart;
import org.openxmlformats.schemas.drawingml.x2006.chart.*;
import org.openxmlformats.schemas.drawingml.x2006.main.STSchemeColorVal;

import java.io.IOException;
import java.io.OutputStream;
import java.math.BigDecimal;
import java.util.List;

/**
 * @author yeehaw
 */
public class LineChartRenderUtil {

    public static void renderChart(XWPFChart chart, LineChartRenderData data) {
        // 获取图表中的 单个图 之所以是数组 可能存在复合图表
        List<XDDFChartData> series = chart.getChartSeries();
        for (XDDFChartData chartData : series) {
            // 折线图
            XDDFLineChartData lineChartData = (XDDFLineChartData) chartData;
            CTLineChart lineChart = (CTLineChart) ReflectUtil.getFieldValue(lineChartData, "chart");
            // 清空
            lineChart.getSerList().clear();
            for (int i = 0; i < data.getLines().size(); i++) {
                LineChartRenderData.LineData lineData = data.getLines().get(i);
                int size = data.getPointX().size();
                // 系列名称 单条线名称 对应  <c:tx></c:tx>
                CTLineSer ctLineSer = lineChart.addNewSer();
                ctLineSer.addNewIdx().setVal(i);
                ctLineSer.addNewOrder().setVal(i);

                CTSerTx tx = ctLineSer.addNewTx();
                String lineRange = new CellRangeAddress(0, 0, i + 1, i + 1).formatAsString("Sheet1", true);
                CTStrRef txCTStrRef = tx.addNewStrRef();
                txCTStrRef.setF(lineRange);
                CTStrData ctStrData = txCTStrRef.addNewStrCache();
                CTStrVal txCTStrVal = ctStrData.addNewPt();
                txCTStrVal.setV(lineData.getTitle());
                txCTStrVal.setIdx(0);

                ctStrData.addNewPtCount().setVal(1);
                // <c:spPr></c:spPr>
                // <c:spPr>
                //      <a:ln w="28575" cap="rnd">
                //        <a:solidFill>
                //          <a:schemeClr val="accent1"/>
                //        </a:solidFill>
                //        <a:round/>
                //      </a:ln>
                //      <a:effectLst/>
                // </c:spPr>
                ctLineSer.addNewSpPr().addNewLn().addNewSolidFill().addNewSchemeClr().setVal(STSchemeColorVal.Enum.forInt(7));
                //     <c:marker>
                //      <c:symbol val="none"/>
                //    </c:marker>
                ctLineSer.addNewMarker().addNewSymbol().setNil();
                //     <c:smooth val="0"/>
                ctLineSer.addNewSmooth().setVal(false);
                // <c:extLst>
                //      <c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">
                //        <c16:uniqueId val="{00000000-2505-4B5B-950D-4F4D615D4F57}"/>
                //      </c:ext>
                //    </c:extLst>
                ctLineSer.addNewExtLst();


                // <c:cat>
                //      <c:strRef>
                //        <c:f>Sheet1!$A$2:$A$3</c:f>
                //        <c:strCache>
                //          <c:ptCount val="2"/>
                //          <c:pt idx="0">
                //            <c:v>类别 1</c:v>
                //          </c:pt>
                //          <c:pt idx="1">
                //            <c:v>类别 2</c:v>
                //          </c:pt>
                //        </c:strCache>
                //      </c:strRef>
                //    </c:cat>
                // <c:cat></c:cat> x 轴 字段名
                CTAxDataSource cat = ctLineSer.addNewCat();
                // <c:val></c:val> y 轴 数据
                CTNumDataSource val = ctLineSer.addNewVal();

                // 渲染 X 轴
                // <c:cat> <c:strRef>   </c:strRef> </c:cat>
                CTStrRef ctStrRef = cat.addNewStrRef();
                // <c:f>Sheet1!$A$2:$A$3</c:f>
                // excel 表格 x 轴 数据范围， 第一列 1 - size
                String xRange = new CellRangeAddress(1, size, 0, 0).formatAsString("Sheet1", true);
                ctStrRef.setF(xRange);
                // X 轴数据
                CTStrData strData = ctStrRef.addNewStrCache();
                // 总数
                strData.addNewPtCount().setVal(size);
                //   <c:pt idx="0">
                //      <c:v>X1</c:v>
                //    </c:pt>
                for (int j = 0; j < size; j++) {
                    CTStrVal ctStrVal = strData.addNewPt();
                    ctStrVal.setIdx(j);
                    ctStrVal.setV(data.getPointX().get(j));
                }
                // 渲染 点位 数据
                // <c:numRef>
                //        <c:f>Sheet1!$B$2:$B$3</c:f>
                //        <c:numCache>
                //          <c:formatCode>General</c:formatCode>
                //          <c:ptCount val="2"/>
                //          <c:pt idx="0">
                //            <c:v>4.3</c:v>
                //          </c:pt>
                //          <c:pt idx="1">
                //            <c:v>2.5</c:v>
                //          </c:pt>
                //        </c:numCache>
                //      </c:numRef>
                CTNumRef ctNumRef = val.addNewNumRef();
                // excel 表格 y 轴 数据范围， 第 i + 1 列 1 - size
                // CellRangeAddress 参数范围 https://blog.csdn.net/aerchi/article/details/7787891
                // CellRangeAddress(起始行号，终止行号， 起始列号，终止列号）
                String yRange = new CellRangeAddress(1, size, i + 1, i + 1).formatAsString("Sheet1", true);
                ctNumRef.setF(yRange);
                CTNumData numData = ctNumRef.addNewNumCache();
                // 总数
                numData.addNewPtCount().setVal(size);
                numData.setFormatCode("General");
                for (int j = 0; j < size; j++) {
                    String key = data.getPointX().get(j);
                    BigDecimal value = lineData.getPoints().get(key);
                    CTNumVal ctNumVal = numData.addNewPt();
                    ctNumVal.setIdx(j);
                    ctNumVal.setV(value.toString());
                }

            }

        }
    }

    public static void renderExcel(XWPFChart chart, LineChartRenderData data) throws IOException {
        XSSFWorkbook wb = new XSSFWorkbook();
        Sheet sheet = wb.createSheet("Sheet1");
        List<LineChartRenderData.LineData> lines = data.getLines();
        //根据数据创建excel第一行标题行
        for (int i = 0; i < lines.size(); i++) {
            if (sheet.getRow(0) == null) {
                sheet.createRow(0).createCell(i + 1).setCellValue(lines.get(i).getTitle());
            } else {
                sheet.getRow(0).createCell(i + 1).setCellValue(lines.get(i).getTitle());
            }
        }
        // 渲染数据
        for (int i = 0; i < data.getPointX().size(); i++) {
            String key = data.getPointX().get(i);
            Row row = sheet.createRow(i + 1);
            row.createCell(0).setCellValue(key);
            for (int j = 0; j < lines.size(); j++) {
                row.createCell(j + 1).setCellValue(lines.get(j).getPoints().get(key).doubleValue());
            }

        }
        chart.setWorkbook(wb);
//        List<POIXMLDocumentPart> pxdList = chart.getRelations();
//        if (pxdList != null && pxdList.size() > 0) {
//            for (int i = 0; i < pxdList.size(); i++) {
//                // 判断为sheet再去进行更新表格数据
//                if (pxdList.get(i).toString().contains("sheet")) {
//                    POIXMLDocumentPart xlsPart = pxdList.get(i);
//                    OutputStream xlsOut = xlsPart.getPackagePart().getOutputStream();
//                    try {
//                        wb.write(xlsOut);
//                        xlsOut.close();
//                        break;
//                    } finally {
//                        if (wb != null) {
//                            wb.close();
//                        }
//                    }
//                }
//            }
//        }
    }

}

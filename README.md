##### apache-poi-word-chart 
使用 POI 操作模板图表
##### 使用
1. 为 [poi-tl](https://github.com/Sayi/poi-tl) 的一个自定义插件
2. 根据word模板，定位折线图图表，尽可能不修改样式，渲染图表
3. [详细描述](https://www.jianshu.com/p/6a60b98effb9)

##### 示例

```
Configure configure = Configure.newBuilder()
                .referencePolicy(
                        new InsertLineChartRefRenderPolicy("###", data))
                .build();

Map<String, Object> params = new HashMap<String, Object>();

XWPFTemplate template = XWPFTemplate
        .compile(templatePath, configure)
        .render(params);
template.writeToFile(after);
```

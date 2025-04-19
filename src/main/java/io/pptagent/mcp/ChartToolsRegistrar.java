package io.pptagent.mcp;

import io.modelcontextprotocol.server.McpServerFeatures;
import io.modelcontextprotocol.spec.McpSchema;
import io.modelcontextprotocol.spec.McpSchema.TextContent;
import io.pptagent.tools.chart.ChartTools;
import io.pptagent.tools.chart.ChartTools.ChartResult;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import reactor.core.publisher.Mono;

/**
 * 图表工具注册类
 */
public class ChartToolsRegistrar {

    /**
     * 创建所有图表工具规范
     */
    public static List<McpServerFeatures.AsyncToolSpecification> createToolSpecifications() {
        List<McpServerFeatures.AsyncToolSpecification> tools = new ArrayList<>();
        
        tools.add(createAddChartToolSpec());
        
        return tools;
    }
    
    /**
     * 创建添加图表工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createAddChartToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {
                "chartType": {
                  "type": "string",
                  "enum": ["COLUMN", "PIE", "LINE"],
                  "description": "图表类型：柱状图(COLUMN)、饼图(PIE)或折线图(LINE)"
                },
                "x": {
                  "type": "number",
                  "minimum": 0,
                  "description": "X坐标位置"
                },
                "y": {
                  "type": "number",
                  "minimum": 0,
                  "description": "Y坐标位置"
                },
                "width": {
                  "type": "number",
                  "minimum": 0,
                  "description": "图表宽度"
                },
                "height": {
                  "type": "number",
                  "minimum": 0,
                  "description": "图表高度"
                },
                "title": {
                  "type": "string",
                  "description": "图表标题"
                },
                "categories": {
                  "type": "array",
                  "items": {"type": "string"},
                  "description": "类别标签列表（对于饼图，这是切片标签）"
                },
                "seriesLabels": {
                  "type": "array",
                  "items": {"type": "string"},
                  "description": "系列标签列表（对于饼图，只使用第一个标签）"
                },
                "seriesData": {
                  "type": "array",
                  "items": {
                    "type": "array",
                    "items": {"type": "number"}
                  },
                  "description": "系列数据，二维数组，第一维是系列，第二维是该系列的值（对于饼图，只使用第一个系列的数据）"
                },
                "seriesColors": {
                  "type": "array",
                  "items": {"type": "string"},
                  "description": "系列颜色，十六进制颜色代码，如'#FF0000'表示红色（对于饼图，这些是切片颜色）"
                },
                "backgroundColor": {
                  "type": "string",
                  "description": "图表背景颜色，十六进制颜色代码，如'#FFFFFF'表示白色"
                },
                "borderColor": {
                  "type": "string",
                  "description": "图表边框颜色，十六进制颜色代码，如'#000000'表示黑色"
                },
                "borderWidth": {
                  "type": "number",
                  "minimum": 0,
                  "description": "图表边框宽度"
                },
                "slideIndex": {
                  "type": "integer",
                  "minimum": 0,
                  "description": "要添加图表的幻灯片索引，从0开始"
                }
              },
              "required": ["chartType", "x", "y", "width", "height", "title", "categories", "seriesLabels", "seriesData", "slideIndex"]
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("addChart", "添加图表到幻灯片", schema),
            (exchange, args) -> {
                try {
                    String chartTypeStr = (String) args.get("chartType");
                    ChartTools.ChartTypeEnum chartType = ChartTools.ChartTypeEnum.valueOf(chartTypeStr);
                    
                    float x = ((Number) args.get("x")).floatValue();
                    float y = ((Number) args.get("y")).floatValue();
                    float width = ((Number) args.get("width")).floatValue();
                    float height = ((Number) args.get("height")).floatValue();
                    String title = (String) args.get("title");
                    @SuppressWarnings("unchecked")
                    List<String> categories = (List<String>) args.get("categories");
                    @SuppressWarnings("unchecked")
                    List<String> seriesLabels = (List<String>) args.get("seriesLabels");
                    @SuppressWarnings("unchecked")
                    List<List<Double>> seriesData = convertToDoubleList((List<List<Number>>) args.get("seriesData"));
                    @SuppressWarnings("unchecked")
                    List<String> seriesColors = args.containsKey("seriesColors") ? 
                        (List<String>) args.get("seriesColors") : null;
                    String backgroundColor = (String) args.get("backgroundColor");
                    String borderColor = (String) args.get("borderColor");
                    float borderWidth = args.containsKey("borderWidth") ? 
                        ((Number) args.get("borderWidth")).floatValue() : 1.0f;
                    int slideIndex = ((Number) args.get("slideIndex")).intValue();
                    
                    // 创建图表参数
                    ChartTools.ChartParams params = ChartTools.ChartParams.builder()
                        .x(x)
                        .y(y)
                        .width(width)
                        .height(height)
                        .title(title)
                        .backgroundColor(backgroundColor)
                        .borderColor(borderColor)
                        .borderWidth(borderWidth)
                        .build();
                    
                    // 创建系列数据列表
                    List<ChartTools.SeriesData> seriesDataList = new ArrayList<>();
                    for (int i = 0; i < seriesLabels.size() && i < seriesData.size(); i++) {
                        String color = (seriesColors != null && i < seriesColors.size()) ? seriesColors.get(i) : null;
                        seriesDataList.add(ChartTools.SeriesData.builder()
                            .name(seriesLabels.get(i))
                            .values(seriesData.get(i))
                            .color(color)
                            .build());
                    }
                    
                    // 调用图表创建方法
                    ChartResult result = ChartTools.createChart(chartType, params, categories, seriesDataList, slideIndex);
                    
                    Map<String, Object> response = new HashMap<>();
                    response.put("success", result.isSuccess());
                    response.put("chartIndex", result.getChartIndex());
                    response.put("message", result.getMessage());
                    
                    // 将结果转为JSON字符串
                    String resultJson = response.toString();
                    
                    // 创建文本内容
                    List<McpSchema.Content> content = List.of(
                        new TextContent(resultJson)
                    );
                    
                    // 使用内容列表创建调用结果
                    return Mono.just(new McpSchema.CallToolResult(content, false));
                } catch (Exception e) {
                    Map<String, Object> errorResponse = new HashMap<>();
                    errorResponse.put("success", false);
                    errorResponse.put("chartIndex", -1);
                    errorResponse.put("message", "添加图表失败: " + e.getMessage());
                    
                    List<McpSchema.Content> errorContent = List.of(
                        new TextContent(errorResponse.toString())
                    );
                    
                    return Mono.just(new McpSchema.CallToolResult(errorContent, false));
                }
            }
        );
    }
    
    /**
     * 将Number类型的二维列表转换为Double类型的二维列表
     */
    private static List<List<Double>> convertToDoubleList(List<List<Number>> input) {
        if (input == null) return null;
        
        List<List<Double>> result = new ArrayList<>();
        for (List<Number> numbers : input) {
            List<Double> doubles = new ArrayList<>();
            for (Number number : numbers) {
                doubles.add(number.doubleValue());
            }
            result.add(doubles);
        }
        return result;
    }
}
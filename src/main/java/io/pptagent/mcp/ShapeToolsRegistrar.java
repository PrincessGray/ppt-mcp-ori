package io.pptagent.mcp;

import io.modelcontextprotocol.server.McpServerFeatures;
import io.modelcontextprotocol.spec.McpSchema;
import io.modelcontextprotocol.spec.McpSchema.TextContent;
import io.pptagent.tools.shape.ShapeTools;
import io.pptagent.tools.shape.ShapeTools.AddShapeResult;
import io.pptagent.tools.shape.ShapeTools.AddLineResult;
import io.pptagent.tools.shape.ShapeTools.ShapeParams;
import io.pptagent.tools.shape.ShapeTools.LineParams;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import reactor.core.publisher.Mono;

/**
 * 形状工具注册类
 */
public class ShapeToolsRegistrar {

    /**
     * 创建所有形状工具规范
     */
    public static List<McpServerFeatures.AsyncToolSpecification> createToolSpecifications() {
        List<McpServerFeatures.AsyncToolSpecification> tools = new ArrayList<>();
        
        //tools.add(createAddShapeToolSpec());
        tools.add(createAddLineToolSpec());
        
        return tools;
    }
    
    /**
     * 创建添加形状工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createAddShapeToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {
                "type": {
                  "type": "string",
                  "enum": ["RECTANGLE", "ELLIPSE", "CIRCLE", "TRIANGLE", "DIAMOND", "STAR", "ARROW"],
                  "description": "形状类型"
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
                  "description": "形状宽度"
                },
                "height": {
                  "type": "number",
                  "minimum": 0,
                  "description": "形状高度"
                },
                "fillColor": {
                  "type": "string",
                  "description": "填充颜色，如#FF0000表示红色"
                },
                "borderColor": {
                  "type": "string",
                  "description": "边框颜色，如#000000表示黑色"
                },
                "borderWidth": {
                  "type": "number",
                  "minimum": 0,
                  "description": "边框宽度"
                },
                "slideIndex": {
                  "type": "integer",
                  "minimum": 0,
                  "description": "要添加形状的幻灯片索引，从0开始"
                }
              },
              "required": ["type", "x", "y", "width", "height", "slideIndex"]
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("addShape", "添加形状到幻灯片", schema),
            (exchange, args) -> {
                String type = (String) args.get("type");
                float x = ((Number) args.get("x")).floatValue();
                float y = ((Number) args.get("y")).floatValue();
                float width = ((Number) args.get("width")).floatValue();
                float height = ((Number) args.get("height")).floatValue();
                String fillColor = (String) args.get("fillColor");
                String borderColor = (String) args.get("borderColor");
                float borderWidth = args.containsKey("borderWidth") ? 
                    ((Number) args.get("borderWidth")).floatValue() : 1.0f;
                int slideIndex = ((Number) args.get("slideIndex")).intValue();
                
                ShapeParams params = new ShapeParams(x, y, width, height, fillColor, borderColor, borderWidth);
                AddShapeResult result = ShapeTools.addShape(type, params, slideIndex);
                
                Map<String, Object> response = new HashMap<>();
                response.put("success", result.isSuccess());
                response.put("shapeIndex", result.getShapeIndex());
                response.put("message", result.getMessage());
                
                // 将结果转为JSON字符串
                String resultJson = response.toString();
                
                // 创建文本内容
                List<McpSchema.Content> content = List.of(
                    new TextContent(resultJson)
                );
                
                // 使用内容列表创建调用结果
                return Mono.just(new McpSchema.CallToolResult(content, false));
            }
        );
    }
    
    /**
     * 创建添加线条工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createAddLineToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {
                "x1": {
                  "type": "number",
                  "minimum": 0,
                  "description": "起点X坐标"
                },
                "y1": {
                  "type": "number",
                  "minimum": 0,
                  "description": "起点Y坐标"
                },
                "x2": {
                  "type": "number",
                  "minimum": 0,
                  "description": "终点X坐标"
                },
                "y2": {
                  "type": "number",
                  "minimum": 0,
                  "description": "终点Y坐标"
                },
                "color": {
                  "type": "string",
                  "description": "线条颜色，如#000000表示黑色"
                },
                "thickness": {
                  "type": "number",
                  "minimum": 0,
                  "description": "线条粗细"
                },
                "slideIndex": {
                  "type": "integer",
                  "minimum": 0,
                  "description": "要添加线条的幻灯片索引，从0开始"
                }
              },
              "required": ["x1", "y1", "x2", "y2", "slideIndex"]
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("addLine", "添加线条到幻灯片", schema),
            (exchange, args) -> {
                float x1 = ((Number) args.get("x1")).floatValue();
                float y1 = ((Number) args.get("y1")).floatValue();
                float x2 = ((Number) args.get("x2")).floatValue();
                float y2 = ((Number) args.get("y2")).floatValue();
                String color = (String) args.get("color");
                float thickness = args.containsKey("thickness") ? 
                    ((Number) args.get("thickness")).floatValue() : 1.0f;
                int slideIndex = ((Number) args.get("slideIndex")).intValue();
                
                LineParams params = new LineParams(x1, y1, x2, y2, color, thickness);
                AddLineResult result = ShapeTools.addLine(params, slideIndex);
                
                Map<String, Object> response = new HashMap<>();
                response.put("success", result.isSuccess());
                response.put("lineIndex", result.getLineIndex());
                response.put("message", result.getMessage());
                
                // 将结果转为JSON字符串
                String resultJson = response.toString();
                
                // 创建文本内容
                List<McpSchema.Content> content = List.of(
                    new TextContent(resultJson)
                );
                
                // 使用内容列表创建调用结果
                return Mono.just(new McpSchema.CallToolResult(content, false));
            }
        );
    }
} 
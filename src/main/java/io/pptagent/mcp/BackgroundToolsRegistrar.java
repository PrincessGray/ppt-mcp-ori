package io.pptagent.mcp;

import io.modelcontextprotocol.server.McpServerFeatures;
import io.modelcontextprotocol.spec.McpSchema;
import io.modelcontextprotocol.spec.McpSchema.TextContent;
import io.pptagent.tools.background.BackgroundTools;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import reactor.core.publisher.Mono;

/**
 * 背景工具注册类
 */
public class BackgroundToolsRegistrar {

    /**
     * 创建所有背景工具规范
     */
    public static List<McpServerFeatures.AsyncToolSpecification> createToolSpecifications() {
        List<McpServerFeatures.AsyncToolSpecification> tools = new ArrayList<>();
        
        //tools.add(createSetBackgroundColorToolSpec());
        tools.add(createSetBackgroundSvgToolSpec());
        
        return tools;
    }
    
    /**
     * 创建设置背景颜色工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createSetBackgroundColorToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {
                "color": {
                  "type": "string",
                  "description": "颜色代码，如#FF0000表示红色"
                },
                "slideIndex": {
                  "type": "integer",
                  "minimum": 0,
                  "description": "要设置背景的幻灯片索引，从0开始"
                }
              },
              "required": ["color", "slideIndex"]
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("setBackgroundColor", "设置幻灯片背景颜色", schema),
            (exchange, args) -> {
                String color = (String) args.get("color");
                int slideIndex = ((Number) args.get("slideIndex")).intValue();
                
                boolean success = BackgroundTools.setBackgroundColor(color, slideIndex);
                
                Map<String, Object> response = new HashMap<>();
                response.put("success", success);
                response.put("message", success ? "背景颜色设置成功" : "背景颜色设置失败");
                
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
     * 创建设置SVG背景工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createSetBackgroundSvgToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {
                "svgContent": {
                  "type": "string",
                  "description": "SVG内容"
                },
                "slideIndex": {
                  "type": "integer",
                  "minimum": 0,
                  "description": "要设置背景的幻灯片索引，从0开始"
                }
              },
              "required": ["svgContent", "slideIndex"]
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("setBackgroundSvg", "设置幻灯片SVG背景", schema),
            (exchange, args) -> {
                String svgContent = (String) args.get("svgContent");
                int slideIndex = ((Number) args.get("slideIndex")).intValue();
                
                boolean success = BackgroundTools.setBackgroundSvg(svgContent, slideIndex);
                
                Map<String, Object> response = new HashMap<>();
                response.put("success", success);
                response.put("message", success ? "SVG背景设置成功" : "SVG背景设置失败");
                
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
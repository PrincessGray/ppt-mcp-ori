package io.pptagent.mcp;

import io.modelcontextprotocol.server.McpServerFeatures;
import io.modelcontextprotocol.spec.McpSchema;
import io.modelcontextprotocol.spec.McpSchema.TextContent;
import io.pptagent.tools.slides.SlideTools;
import io.pptagent.tools.slides.SlideTools.AddSlideResult;
import io.pptagent.tools.slides.SlideTools.SelectSlideResult;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import reactor.core.publisher.Mono;

/**
 * 幻灯片工具注册类
 */
public class SlideToolsRegistrar {

    /**
     * 创建所有幻灯片工具规范
     */
    public static List<McpServerFeatures.AsyncToolSpecification> createToolSpecifications() {
        List<McpServerFeatures.AsyncToolSpecification> tools = new ArrayList<>();
        
        tools.add(createAddSlideToolSpec());
        //tools.add(createSelectSlideToolSpec());
        
        return tools;
    }
    
    /**
     * 创建添加幻灯片工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createAddSlideToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {
                "layoutType": {
                  "type": "string",
                  "enum": ["BLANK", "TITLE", "TITLEBODY", "TITLEONLY"],
                  "description": "幻灯片布局类型：BLANK(空白)、TITLE(标题)、TITLEBODY(标题和内容)、TITLEONLY(仅标题)"
                }
              },
              "required": ["layoutType"]
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("addSlide", "添加一个新的幻灯片", schema),
            (exchange, args) -> {
                String layoutType = (String) args.get("layoutType");
                
                AddSlideResult result = SlideTools.addSlideEnhanced(layoutType);
                
                Map<String, Object> response = new HashMap<>();
                response.put("success", result.isSuccess());
                response.put("slideIndex", result.getSlideIndex());
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
     * 创建选择幻灯片工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createSelectSlideToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {
                "slideIndex": {
                  "type": "integer",
                  "minimum": 0,
                  "description": "要选择的幻灯片索引，从0开始"
                }
              },
              "required": ["slideIndex"]
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("selectSlide", "选择当前操作的幻灯片", schema),
            (exchange, args) -> {
                int slideIndex = ((Number) args.get("slideIndex")).intValue();
                
                SelectSlideResult result = SlideTools.selectSlideEnhanced(slideIndex);
                
                Map<String, Object> response = new HashMap<>();
                response.put("success", result.isSuccess());
                response.put("slideIndex", result.getSlideIndex());
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
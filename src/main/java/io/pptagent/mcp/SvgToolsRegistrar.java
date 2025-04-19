package io.pptagent.mcp;

import io.modelcontextprotocol.server.McpServerFeatures;
import io.modelcontextprotocol.spec.McpSchema;
import io.modelcontextprotocol.spec.McpSchema.TextContent;
import io.pptagent.tools.svg.SvgTools;
import io.pptagent.tools.svg.SvgTools.SvgParams;
import io.pptagent.tools.svg.SvgTools.AddSvgResult;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import reactor.core.publisher.Mono;

/**
 * SVG工具注册类
 */
public class SvgToolsRegistrar {

    /**
     * 创建所有SVG工具规范
     */
    public static List<McpServerFeatures.AsyncToolSpecification> createToolSpecifications() {
        List<McpServerFeatures.AsyncToolSpecification> tools = new ArrayList<>();
        
        tools.add(createAddSvgImageToolSpec());
        
        return tools;
    }
    
    /**
     * 创建添加SVG图像工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createAddSvgImageToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {
                "svgContent": {
                  "type": "string",
                  "description": "SVG图像内容"
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
                  "description": "图像宽度"
                },
                "height": {
                  "type": "number",
                  "minimum": 0,
                  "description": "图像高度"
                },
                "slideIndex": {
                  "type": "integer",
                  "minimum": 0,
                  "description": "要添加SVG图像的幻灯片索引，从0开始"
                }
              },
              "required": ["svgContent", "x", "y", "width", "height", "slideIndex"]
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("addSvgImage", "添加SVG图像到幻灯片", schema),
            (exchange, args) -> {
                String svgContent = (String) args.get("svgContent");
                float x = ((Number) args.get("x")).floatValue();
                float y = ((Number) args.get("y")).floatValue();
                float width = ((Number) args.get("width")).floatValue();
                float height = ((Number) args.get("height")).floatValue();
                int slideIndex = ((Number) args.get("slideIndex")).intValue();
                
                SvgParams params = new SvgParams(svgContent, x, y, width, height);
                AddSvgResult result = SvgTools.addSvgImage(params, slideIndex);
                
                Map<String, Object> response = new HashMap<>();
                response.put("success", result.success());
                response.put("imageIndex", result.imageIndex());
                response.put("message", result.message());
                
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
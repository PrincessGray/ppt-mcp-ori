package io.pptagent.mcp;

import io.modelcontextprotocol.server.McpServerFeatures;
import io.modelcontextprotocol.spec.McpSchema;
import io.modelcontextprotocol.spec.McpSchema.TextContent;
import io.pptagent.tools.info.InfoTools;
import io.pptagent.tools.info.InfoTools.ShapeInfo;
import io.pptagent.tools.PresentationManager;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import reactor.core.publisher.Mono;
import com.aspose.slides.Presentation;

/**
 * 信息工具注册类
 */
public class InfoToolsRegistrar {

    /**
     * 创建所有信息工具规范
     */
    public static List<McpServerFeatures.AsyncToolSpecification> createToolSpecifications() {
        List<McpServerFeatures.AsyncToolSpecification> tools = new ArrayList<>();
        
        tools.add(createGetShapesInfoToolSpec());
        
        return tools;
    }
    
    /**
     * 创建获取幻灯片数量工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createGetSlideCountToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {}
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("getSlideCount", "获取当前演示文稿的幻灯片数量", schema),
            (exchange, args) -> {
                Presentation pres = PresentationManager.getInstance().getPresentation();
                int slideCount = InfoTools.getSlideCount(pres);
                
                Map<String, Object> response = new HashMap<>();
                response.put("slideCount", slideCount);
                
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
     * 创建获取形状信息工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createGetShapesInfoToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {
                "slideIndex": {
                  "type": "integer",
                  "minimum": 0,
                  "description": "要获取形状信息的幻灯片索引，从0开始"
                }
              },
              "required": ["slideIndex"]
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("getShapesInfo", "获取指定幻灯片中所有形状的信息", schema),
            (exchange, args) -> {
                int slideIndex = ((Number) args.get("slideIndex")).intValue();
                
                Presentation pres = PresentationManager.getInstance().getPresentation();
                List<ShapeInfo> shapesInfo = InfoTools.getShapesInfo(pres, slideIndex);
                
                // 构建响应数据
                Map<String, Object> response = new HashMap<>();
                response.put("slideIndex", slideIndex);
                response.put("shapeCount", shapesInfo.size());
                
                List<Map<String, Object>> shapesData = new ArrayList<>();
                for (ShapeInfo shape : shapesInfo) {
                    Map<String, Object> shapeData = new HashMap<>();
                    shapeData.put("shapeIndex", shape.getShapeIndex());
                    shapeData.put("shapeType", shape.getShapeType());
                    shapeData.put("x", shape.getX());
                    shapeData.put("y", shape.getY());
                    shapeData.put("width", shape.getWidth());
                    shapeData.put("height", shape.getHeight());
                    shapeData.put("hasTextFrame", shape.isHasTextFrame());
                    
                    // 只有当有文本框架时才包含文本内容
                    if (shape.isHasTextFrame()) {
                        shapeData.put("textContent", shape.getTextContent());
                    }
                    
                    shapesData.add(shapeData);
                }
                
                response.put("shapes", shapesData);
                
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
package io.pptagent.mcp;

import io.modelcontextprotocol.server.McpServerFeatures;
import io.modelcontextprotocol.spec.McpSchema;
import io.modelcontextprotocol.spec.McpSchema.TextContent;
import io.pptagent.tools.media.PictureTools;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import reactor.core.publisher.Mono;

/**
 * 图片工具注册类
 */
public class PictureToolsRegistrar {

    /**
     * 创建所有图片工具规范
     */
    public static List<McpServerFeatures.AsyncToolSpecification> createToolSpecifications() {
        List<McpServerFeatures.AsyncToolSpecification> tools = new ArrayList<>();
        
        tools.add(createAddPictureFrameToolSpec());
        
        return tools;
    }
    
    /**
     * 创建添加图片框工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createAddPictureFrameToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {
                "imagePath": {
                  "type": "string",
                  "description": "图片文件路径"
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
                  "description": "图片框宽度"
                },
                "height": {
                  "type": "number",
                  "minimum": 0,
                  "description": "图片框高度，宽高比“必须”和原图片宽高比例一致"
                },
                "slideIndex": {
                  "type": "integer",
                  "minimum": 0,
                  "description": "要添加图片的幻灯片索引，从0开始"
                }
              },
              "required": ["imagePath", "x", "y", "width", "height", "slideIndex"]
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("addPictureFrame", "添加图片框到幻灯片，保持图片原始比例", schema),
            (exchange, args) -> {
                String imagePath = (String) args.get("imagePath");
                float x = ((Number) args.get("x")).floatValue();
                float y = ((Number) args.get("y")).floatValue();
                float width = ((Number) args.get("width")).floatValue();
                float height = ((Number) args.get("height")).floatValue();
                
                // 幻灯片索引为可选参数，默认为0
                int slideIndex = args.containsKey("slideIndex") ? 
                                ((Number) args.get("slideIndex")).intValue() : 0;
                
                // 调用图片工具添加图片框
                int frameIndex = PictureTools.addPictureFrameWithAspectRatio(
                    imagePath, x, y, width, height, slideIndex);
                
                // 构建响应
                Map<String, Object> response = new HashMap<>();
                response.put("success", frameIndex >= 0);
                response.put("frameIndex", frameIndex);
                response.put("message", frameIndex >= 0 ? "图片框添加成功" : "图片框添加失败");
                
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
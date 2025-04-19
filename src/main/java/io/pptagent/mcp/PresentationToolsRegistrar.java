package io.pptagent.mcp;

import io.modelcontextprotocol.server.McpServerFeatures;
import io.modelcontextprotocol.spec.McpSchema;
import io.modelcontextprotocol.spec.McpSchema.TextContent;
import io.pptagent.tools.base.BaseTools;
import io.pptagent.tools.base.BaseTools.SaveResult;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import reactor.core.publisher.Mono;

/**
 * 演示文稿基础工具注册类
 */
public class PresentationToolsRegistrar {

    /**
     * 创建所有演示文稿工具规范
     */
    public static List<McpServerFeatures.AsyncToolSpecification> createToolSpecifications() {
        List<McpServerFeatures.AsyncToolSpecification> tools = new ArrayList<>();
        
        tools.add(createPresentationToolSpec());
        tools.add(savePresentationToolSpec());
        
        return tools;
    }
    
    /**
     * 创建演示文稿工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createPresentationToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {}
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("createPresentation", "创建一个新的空白演示文稿", schema),
            (exchange, args) -> {
                boolean success = BaseTools.createPresentation();
                
                Map<String, Object> result = new HashMap<>();
                result.put("success", success);
                result.put("message", success ? "演示文稿创建成功" : "演示文稿创建失败");
                
                // 将结果转为JSON字符串
                String resultJson = result.toString();
                
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
     * 保存演示文稿工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification savePresentationToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {
                "filePath": {
                  "type": "string",
                  "description": "保存文件的完整路径，包括文件名和扩展名"
                },
                "format": {
                  "type": "string",
                  "enum": ["PPTX", "PPT", "PDF"],
                  "description": "文件格式，支持PPTX、PPT或PDF"
                }
              },
              "required": ["filePath", "format"]
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("savePresentation", "保存演示文稿到指定路径", schema),
            (exchange, args) -> {
                String filePath = (String) args.get("filePath");
                String format = (String) args.get("format");
                
                SaveResult result = BaseTools.savePresentationEnhanced(filePath, format);
                
                Map<String, Object> response = new HashMap<>();
                response.put("success", result.isSuccess());
                response.put("message", result.getMessage());
                response.put("path", result.getPath());
                
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
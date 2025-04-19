package io.pptagent.mcp;

import io.modelcontextprotocol.server.McpServerFeatures;
import io.modelcontextprotocol.spec.McpSchema;
import io.modelcontextprotocol.spec.McpSchema.TextContent;
import io.pptagent.tools.animation.AnimationTools;
import io.pptagent.tools.animation.AnimationTools.AddAnimationResult;
import io.pptagent.tools.animation.AnimationTools.AnimationParams;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import reactor.core.publisher.Mono;

/**
 * 动画工具注册类
 */
public class AnimationToolsRegistrar {
    
    /**
     * 创建所有动画工具规范
     */
    public static List<McpServerFeatures.AsyncToolSpecification> createToolSpecifications() {
        List<McpServerFeatures.AsyncToolSpecification> tools = new ArrayList<>();
        
        tools.add(createAddAnimationToolSpec());
        //tools.add(createAddParagraphAnimationToolSpec());
        
        return tools;
    }
    
    /**
     * 创建添加形状动画工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createAddAnimationToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {
                "shapeIndex": {
                  "type": "integer",
                  "minimum": 0,
                  "description": "要添加动画的形状索引"
                },
                "effectType": {
                  "type": "string",
                  "enum": ["APPEAR", "FADE", "FLOAT", "BOUNCE", "WIPE", "FLY"],
                  "description": "动画效果类型"
                },
                "effectSubtype": {
                  "type": "string",
                  "enum": ["NONE(APPEAR,FADE,FLOAT,BOUNCE只能使用NONE)", "LEFT", "RIGHT", "TOP", "BOTTOM", "IN", "OUT"],
                  "description": "动画效果子类型"
                },
                "triggerType": {
                  "type": "string",
                  "enum": ["ONCLICK"],
                  "description": "动画触发类型"
                },
                "slideIndex": {
                  "type": "integer",
                  "minimum": 0,
                  "description": "幻灯片索引，从0开始"
                }
              },
              "required": ["shapeIndex", "effectType", "slideIndex"]
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("addAnimation", "为形状添加动画效果", schema),
            (exchange, args) -> {
                int shapeIndex = ((Number) args.get("shapeIndex")).intValue();
                String effectType = (String) args.get("effectType");
                String effectSubtype = (String) args.get("effectSubtype");
                String triggerType = (String) args.get("triggerType");
                int slideIndex = ((Number) args.get("slideIndex")).intValue();
                
                AnimationParams params = AnimationParams.builder()
                    .effectType(effectType)
                    .effectSubtype(effectSubtype)
                    .triggerType(triggerType)
                    .build();
                
                AddAnimationResult result = AnimationTools.addAnimation(shapeIndex, params, slideIndex);
                
                Map<String, Object> response = new HashMap<>();
                response.put("success", result.isSuccess());
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
     * 创建添加段落动画工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createAddParagraphAnimationToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {
                "shapeIndex": {
                  "type": "integer",
                  "minimum": 0,
                  "description": "包含段落的形状索引"
                },
                "paragraphIndex": {
                  "type": "integer",
                  "minimum": 1,
                  "description": "段落索引，从1开始"
                },
                "effectType": {
                  "type": "string",
                  "enum": ["APPEAR", "FADE", "FLOAT", "BOUNCE", "WIPE", "FLY"],
                  "description": "动画效果类型"
                },
                "effectSubtype": {
                  "type": "string",
                  "enum": ["NONE", "LEFT", "RIGHT", "TOP", "BOTTOM", "IN", "OUT"],
                  "description": "动画效果子类型"
                },
                "triggerType": {
                  "type": "string",
                  "enum": ["ONCLICK", "WITHPREVIOUS", "AFTERPREVIOUS"],
                  "description": "动画触发类型"
                },
                "slideIndex": {
                  "type": "integer",
                  "minimum": 0,
                  "description": "幻灯片索引，从0开始"
                }
              },
              "required": ["shapeIndex", "paragraphIndex", "effectType", "slideIndex"]
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("addParagraphAnimation", "为形状中的段落添加动画效果", schema),
            (exchange, args) -> {
                int shapeIndex = ((Number) args.get("shapeIndex")).intValue();
                int paragraphIndex = ((Number) args.get("paragraphIndex")).intValue();
                String effectType = (String) args.get("effectType");
                String effectSubtype = (String) args.get("effectSubtype");
                String triggerType = (String) args.get("triggerType");
                int slideIndex = ((Number) args.get("slideIndex")).intValue();
                
                AnimationParams params = AnimationParams.builder()
                    .effectType(effectType)
                    .effectSubtype(effectSubtype)
                    .triggerType(triggerType)
                    .build();
                
                AddAnimationResult result = AnimationTools.addParagraphAnimation(
                    shapeIndex, paragraphIndex, params, slideIndex);
                
                Map<String, Object> response = new HashMap<>();
                response.put("success", result.isSuccess());
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
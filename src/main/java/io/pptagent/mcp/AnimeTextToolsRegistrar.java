package io.pptagent.mcp;

import io.modelcontextprotocol.server.McpServerFeatures;
import io.modelcontextprotocol.spec.McpSchema;
import io.modelcontextprotocol.spec.McpSchema.TextContent;
import io.pptagent.tools.text.TextTools;
import io.pptagent.tools.text.TextTools.AddTextBoxResult;
import io.pptagent.tools.text.TextTools.SetFormattedTextResult;
import io.pptagent.tools.text.TextTools.TextBoxParams;

import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import reactor.core.publisher.Mono;

/**
 * 文本工具注册类
 */
public class AnimeTextToolsRegistrar {

    /**
     * 创建所有文本工具规范
     */
    public static List<McpServerFeatures.AsyncToolSpecification> createToolSpecifications() {
        List<McpServerFeatures.AsyncToolSpecification> tools = new ArrayList<>();
        
        tools.add(createAddTextBoxToolSpec());
        tools.add(createSetFormattedTextToolSpec());
        
        return tools;
    }
    
    /**
     * 创建添加文本框工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createAddTextBoxToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {
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
                  "description": "文本框宽度"
                },
                "height": {
                  "type": "number",
                  "minimum": 0,
                  "description": "文本框高度"
                },
                "backgroundColor": {
                  "type": "string",
                  "description": "背景颜色，如#FFFFFF表示白色"
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
                  "description": "要添加文本框的幻灯片索引，从0开始"
                }
              },
              "required": ["x", "y", "width", "height", "slideIndex"]
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("addTextBox", "添加文本框到幻灯片", schema),
            (exchange, args) -> {
                float x = ((Number) args.get("x")).floatValue();
                float y = ((Number) args.get("y")).floatValue();
                float width = ((Number) args.get("width")).floatValue();
                float height = ((Number) args.get("height")).floatValue();
                String backgroundColor = (String) args.get("backgroundColor");
                String borderColor = (String) args.get("borderColor");
                float borderWidth = args.containsKey("borderWidth") ? 
                    ((Number) args.get("borderWidth")).floatValue() : 1.0f;
                int slideIndex = ((Number) args.get("slideIndex")).intValue();
                
                TextBoxParams params = new TextBoxParams(x, y, width, height, backgroundColor, borderColor, borderWidth);
                AddTextBoxResult result = TextTools.addTextBox(params, slideIndex);
                
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
     * 创建设置格式化文本工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createSetFormattedTextToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {
                "shapeIndex": {
                  "type": "integer",
                  "minimum": 0,
                  "description": "要设置文本的形状索引"
                },
                "formattedText": {
                  "type": "array",
                  "items": {
                    "type": "object",
                    "properties": {
                      "text": {
                        "type": "string",
                        "description": "文本内容"
                      },
                      "fontName": {
                        "type": "string",
                        "description": "字体名称"
                      },
                      "fontSize": {
                        "type": "number",
                        "description": "字体大小"
                      },
                      "bold": {
                        "type": "boolean",
                        "description": "是否粗体"
                      },
                      "italic": {
                        "type": "boolean",
                        "description": "是否斜体"
                      },
                      "color": {
                        "type": "string",
                        "description": "文本颜色，如#000000表示黑色"
                      },
                      "animation": {
                        "type": "object",
                        "properties": {
                          "effectType": {
                            "type": "string",
                            "description": "动画效果类型，仅支持'APPEAR'、'FADE'、'FLOAT'、'BOUNCE'、'WIPE'、'GROW'、'SPIN'"
                          },
                          "effectSubtype": {
                            "type": "string",
                            "description": "动画效果子类型，仅支持'LEFT'、'RIGHT'、'TOP'、'BOTTOM'、'IN'、'OUT'"
                          },
                          "triggerType": {
                            "type": "string",
                            "description": "动画触发类型，仅支持'ONCLICK'、'WITHPREVIOUS'、'AFTERPREVIOUS'"
                          }
                        },
                        "description": "动画参数"
                      }
                    }
                  },
                  "description": "格式化文本数组"
                },
                "slideIndex": {
                  "type": "integer",
                  "minimum": 0,
                  "description": "幻灯片索引，从0开始"
                }
              },
              "required": ["shapeIndex", "formattedText", "slideIndex"]
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("setFormattedText", "设置形状的格式化文本", schema),
            (exchange, args) -> {
                int shapeIndex = ((Number) args.get("shapeIndex")).intValue();
                List<Map<String, Object>> formattedText = (List<Map<String, Object>>) args.get("formattedText");
                int slideIndex = ((Number) args.get("slideIndex")).intValue();
                
                SetFormattedTextResult result = TextTools.setFormattedText(shapeIndex, formattedText, slideIndex);
                
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
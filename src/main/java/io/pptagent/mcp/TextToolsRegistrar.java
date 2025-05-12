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
public class TextToolsRegistrar {
    
    /**
     * 创建所有文本工具规范
     */
    public static List<McpServerFeatures.AsyncToolSpecification> createToolSpecifications() {
        List<McpServerFeatures.AsyncToolSpecification> tools = new ArrayList<>();
        
        //tools.add(createAddTextBoxToolSpec());
        tools.add(createSetFormattedTextToolSpec());
        tools.add(createAddMultipleTextBoxesToolSpec());
        
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
                        "description": "文本颜色"
                      }
                    },
                    "required": ["text"]
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
    
    /**
     * 创建批量添加文本框工具规范
     */
    private static McpServerFeatures.AsyncToolSpecification createAddMultipleTextBoxesToolSpec() {
        String schema = """
            {
              "type": "object",
              "properties": {
                "textBoxes": {
                  "type": "array",
                  "items": {
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
                      "text": {
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
                              "description": "文本颜色"
                            }
                          },
                          "required": ["text"]
                        },
                        "description": "格式化文本数组"
                      }
                    },
                    "required": ["x", "y", "width", "height"]
                  },
                  "description": "要添加的文本框列表"
                },
                "slideIndex": {
                  "type": "integer",
                  "minimum": 0,
                  "description": "要添加文本框的幻灯片索引，从0开始"
                }
              },
              "required": ["textBoxes", "slideIndex"]
            }
            """;
        
        return new McpServerFeatures.AsyncToolSpecification(
            new McpSchema.Tool("addMultipleTextBoxes", "批量添加文本框到幻灯片", schema),
            (exchange, args) -> {
                // 获取参数
                List<Map<String, Object>> textBoxes = (List<Map<String, Object>>) args.get("textBoxes");
                int slideIndex = ((Number) args.get("slideIndex")).intValue();
                
                // 用于存储批量操作的结果
                List<Map<String, Object>> results = new ArrayList<>();
                
                // 依次处理每个文本框
                for (Map<String, Object> textBoxData : textBoxes) {
                    try {
                        // 提取文本框参数
                        float x = ((Number) textBoxData.get("x")).floatValue();
                        float y = ((Number) textBoxData.get("y")).floatValue();
                        float width = ((Number) textBoxData.get("width")).floatValue();
                        float height = ((Number) textBoxData.get("height")).floatValue();
                        String backgroundColor = (String) textBoxData.get("backgroundColor");
                        String borderColor = (String) textBoxData.get("borderColor");
                        float borderWidth = textBoxData.containsKey("borderWidth") ? 
                            ((Number) textBoxData.get("borderWidth")).floatValue() : 1.0f;
                        
                        // 创建文本框
                        TextBoxParams params = new TextBoxParams(x, y, width, height, backgroundColor, borderColor, borderWidth);
                        AddTextBoxResult addResult = TextTools.addTextBox(params, slideIndex);
                        
                        // 如果成功创建文本框且有文本内容，则设置文本
                        if (addResult.isSuccess() && textBoxData.containsKey("text")) {
                            List<Map<String, Object>> formattedText = (List<Map<String, Object>>) textBoxData.get("text");
                            if (formattedText != null && !formattedText.isEmpty()) {
                                SetFormattedTextResult textResult = TextTools.setFormattedText(addResult.getShapeIndex(), formattedText, slideIndex);
                                
                                // 记录完整结果
                                Map<String, Object> result = new HashMap<>();
                                result.put("success", addResult.isSuccess() && textResult.isSuccess());
                                result.put("shapeIndex", addResult.getShapeIndex());
                                result.put("message", textResult.isSuccess() ? 
                                    "文本框创建并设置文本成功" : "文本框创建成功但设置文本失败: " + textResult.getMessage());
                                
                                results.add(result);
                            } else {
                                // 只有文本框，没有实际文本内容
                                Map<String, Object> result = new HashMap<>();
                                result.put("success", addResult.isSuccess());
                                result.put("shapeIndex", addResult.getShapeIndex());
                                result.put("message", "文本框创建成功");
                                
                                results.add(result);
                            }
                        } else {
                            // 记录只创建文本框的结果
                            Map<String, Object> result = new HashMap<>();
                            result.put("success", addResult.isSuccess());
                            result.put("shapeIndex", addResult.getShapeIndex());
                            result.put("message", addResult.getMessage());
                            
                            results.add(result);
                        }
                    } catch (Exception e) {
                        // 处理单个文本框创建过程中的异常
                        Map<String, Object> result = new HashMap<>();
                        result.put("success", false);
                        result.put("message", "创建文本框时出错: " + e.getMessage());
                        
                        results.add(result);
                    }
                }
                
                // 构建整体响应
                Map<String, Object> response = new HashMap<>();
                response.put("totalCount", textBoxes.size());
                response.put("successCount", results.stream().filter(r -> (boolean)r.get("success")).count());
                response.put("results", results);
                
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
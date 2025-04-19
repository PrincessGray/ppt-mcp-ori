package io.pptagent.tools.text;

import java.awt.Color;
import java.util.List;
import java.util.Map;
import java.util.HashMap;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.aspose.slides.FillType;
import com.aspose.slides.IAutoShape;
import com.aspose.slides.ISlide;
import com.aspose.slides.ITextFrame;
import com.aspose.slides.NullableBool;
import com.aspose.slides.Paragraph;
import com.aspose.slides.Portion;
import com.aspose.slides.Presentation;
import com.aspose.slides.ShapeType;
import com.aspose.slides.FontData;
import com.aspose.slides.IEffect;
import com.aspose.slides.EffectType;
import com.aspose.slides.EffectSubtype;
import com.aspose.slides.EffectTriggerType;
import com.aspose.slides.IParagraph;
import io.pptagent.tools.PresentationManager;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 带动画文本相关工具函数
 */
public final class AnimeTextTools {
    private static final Logger LOGGER = Logger.getLogger(AnimeTextTools.class.getName());
    
    // 效果类型映射
    private static final Map<String, Integer> EFFECT_TYPE_MAP = Map.of(
        "APPEAR", EffectType.Appear,    // 出现
        "FADE", EffectType.Fade,        // 淡入淡出
        "FLOAT", EffectType.Float,      // 浮动
        "BOUNCE", EffectType.Bounce,    // 弹跳
        "WIPE", EffectType.Wipe,         // 擦除
        "FLY", EffectType.Fly           // 飞入/飞出
    );
    
    // 效果子类型映射
    private static final Map<String, Integer> EFFECT_SUBTYPE_MAP = Map.of(
        "NONE", EffectSubtype.None,
        "LEFT", EffectSubtype.Left,
        "RIGHT", EffectSubtype.Right,
        "TOP", EffectSubtype.Top,
        "BOTTOM", EffectSubtype.Bottom,
        "IN", EffectSubtype.In,
        "OUT", EffectSubtype.Out
    );
    
    // 触发类型映射
    private static final Map<String, Integer> TRIGGER_TYPE_MAP = Map.of(
        "ONCLICK", EffectTriggerType.OnClick,
        //"WITHPREVIOUS", EffectTriggerType.WithPrevious,
        "AFTERPREVIOUS", EffectTriggerType.AfterPrevious
    );
    
    private AnimeTextTools() {
        // 私有构造函数防止实例化
    }
    
    /**
     * 表示文本框参数的类
     */
    @Getter
    @AllArgsConstructor
    public static class TextBoxParams {
        private final float x;
        private final float y;
        private final float width;
        private final float height;
        private final String backgroundColor;
        private final String borderColor;
        private final float borderWidth;
    }
    
    /**
     * 表示添加文本框结果的类
     */
    @Getter
    @AllArgsConstructor
    public static class AddTextBoxResult {
        private final boolean success;
        private final int shapeIndex;
        private final String message;
    }
    
    /**
     * 添加文本框
     * 
     * @param x X坐标
     * @param y Y坐标
     * @param width 宽度
     * @param height 高度
     * @param backgroundColor 背景色
     * @param borderColor 边框颜色
     * @param borderWidth 边框宽度
     * @param slideIndex 幻灯片索引
     * @return 添加文本框结果
     */
    public static AddTextBoxResult addTextBox(float x, float y, float width, float height, 
                               String backgroundColor, String borderColor, float borderWidth, int slideIndex) {
        TextBoxParams params = new TextBoxParams(x, y, width, height, backgroundColor, borderColor, borderWidth);
        return addTextBox(params, slideIndex);
    }
    
    /**
     * 添加文本框
     * 
     * @param params 文本框参数
     * @param slideIndex 幻灯片索引
     * @return 添加文本框结果
     */
    public static AddTextBoxResult addTextBox(TextBoxParams params, int slideIndex) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                return new AddTextBoxResult(false, -1, "没有活动的演示文稿");
            }
            
            ISlide slide = pres.getSlides().get_Item(slideIndex);
            
            // 创建自动形状作为文本框
            IAutoShape textBox = slide.getShapes().addAutoShape(
                ShapeType.Rectangle, 
                params.getX(), 
                params.getY(), 
                params.getWidth(), 
                params.getHeight()
            );
            
            // 设置背景颜色
            if (params.getBackgroundColor() != null && !params.getBackgroundColor().isEmpty()) {
                textBox.getFillFormat().setFillType(FillType.Solid);
                Color bgColor = Color.decode(params.getBackgroundColor());
                textBox.getFillFormat().getSolidFillColor().setColor(bgColor);
            } else {
                textBox.getFillFormat().setFillType(FillType.NoFill);
            }
            
            // 设置边框
            if (params.getBorderColor() != null && !params.getBorderColor().isEmpty()) {
                textBox.getLineFormat().getFillFormat().setFillType(FillType.Solid);
                Color borderColor = Color.decode(params.getBorderColor());
                textBox.getLineFormat().getFillFormat().getSolidFillColor().setColor(borderColor);
                textBox.getLineFormat().setWidth(params.getBorderWidth());
            } else {
                textBox.getLineFormat().getFillFormat().setFillType(FillType.NoFill);
            }
            
            // 添加空文本框架
            textBox.addTextFrame("");
            
            // 获取形状索引
            int shapeIndex = slide.getShapes().indexOf(textBox);
            return new AddTextBoxResult(true, shapeIndex, "文本框添加成功");
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "添加文本框失败", e);
            return new AddTextBoxResult(false, -1, "添加文本框失败: " + e.getMessage());
        }
    }
    
    /**
     * 表示格式化文本操作结果的类
     */
    @Getter
    @AllArgsConstructor
    public static class SetFormattedTextResult {
        private final boolean success;
        private final String message;
    }
    
    /**
     * 设置文本框中的多格式文本
     * 完全重置文本框中的内容
     * 
     * @param shapeIndex 形状索引
     * @param formattedText 格式化文本结构数组，可以包含animation属性
     * @param slideIndex 幻灯片索引
     * @return 设置格式化文本结果
     */
    public static SetFormattedTextResult setFormattedText(int shapeIndex, List<Map<String, Object>> formattedText, int slideIndex) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                return new SetFormattedTextResult(false, "没有活动的演示文稿");
            }
            
            ISlide slide = pres.getSlides().get_Item(slideIndex);
            
            if (shapeIndex < 0 || shapeIndex >= slide.getShapes().size()) {
                return new SetFormattedTextResult(false, "无效的形状索引");
            }
            
            if (!(slide.getShapes().get_Item(shapeIndex) instanceof IAutoShape shape)) {
                return new SetFormattedTextResult(false, "指定的形状不是自动形状");
            }

            // 确保形状有文本框架
            if (!shape.isTextBox()) {
                shape.addTextFrame("");
            }
            
            ITextFrame textFrame = shape.getTextFrame();
            
            // 完全清除现有段落
            textFrame.getParagraphs().clear();
            
            // 创建新段落
            Paragraph para = new Paragraph();
            
            // 提取动画参数
            Map<String, String> animationParams = null;
            
            // 为每个格式化文本部分创建Portion并设置格式
            for (Map<String, Object> textPart : formattedText) {
                // 检查是否包含动画属性
                if (textPart.containsKey("animation")) {
                    // 如果包含动画属性，提取动画参数
                    @SuppressWarnings("unchecked")
                    Map<String, Object> animation = (Map<String, Object>) textPart.get("animation");
                    
                    animationParams = new HashMap<>();
                    
                    if (animation.containsKey("effectType")) {
                        animationParams.put("effectType", animation.get("effectType").toString());
                    }
                    if (animation.containsKey("effectSubtype")) {
                        animationParams.put("effectSubtype", animation.get("effectSubtype").toString());
                    }
                    if (animation.containsKey("triggerType")) {
                        animationParams.put("triggerType", animation.get("triggerType").toString());
                    }
                    
                    // 动画参数处理完成，继续下一个文本部分
                    continue;
                }
                
                Portion portion = new Portion();
                
                // 设置文本内容
                String text = (String) textPart.get("text");
                portion.setText(text);
                
                // 应用字体设置
                if (textPart.containsKey("fontName")) {
                    portion.getPortionFormat().setLatinFont(new FontData((String) textPart.get("fontName")));
                }
                
                // 应用字体大小
                if (textPart.containsKey("fontSize")) {
                    Number fontSize = (Number) textPart.get("fontSize");
                    portion.getPortionFormat().setFontHeight((float)fontSize.doubleValue());
                }
                
                // 应用粗体
                if (textPart.containsKey("bold") && (boolean) textPart.get("bold")) {
                    portion.getPortionFormat().setFontBold(NullableBool.True);
                }
                
                // 应用斜体
                if (textPart.containsKey("italic") && (boolean) textPart.get("italic")) {
                    portion.getPortionFormat().setFontItalic(NullableBool.True);
                }
                
                // 应用颜色
                if (textPart.containsKey("color")) {
                    portion.getPortionFormat().getFillFormat().setFillType(FillType.Solid);
                    Color color = Color.decode((String) textPart.get("color"));
                    portion.getPortionFormat().getFillFormat().getSolidFillColor().setColor(color);
                }
                
                // 将Portion添加到段落
                para.getPortions().add(portion);
            }
            
            // 将段落添加到文本框
            textFrame.getParagraphs().add(para);
            
            // 如果有动画参数，为段落添加动画
            if (animationParams != null) {
                try {
                    // 获取刚添加的段落 - 段落索引从1开始
                    IParagraph paragraph = textFrame.getParagraphs().get_Item(1);
                    
                    // 获取动画参数
                    String effectType = animationParams.get("effectType");
                    String effectSubtype = animationParams.get("effectSubtype");
                    String triggerType = animationParams.get("triggerType");
                    
                    // 将字符串映射为枚举值
                    int effectTypeValue = getEffectTypeValue(effectType);
                    int effectSubtypeValue = getEffectSubtypeValue(effectSubtype);
                    int triggerTypeValue = getTriggerTypeValue(triggerType);
                    
                    // 添加动画效果
                    IEffect effect = slide.getTimeline().getMainSequence().addEffect(
                        paragraph, effectTypeValue, effectSubtypeValue, triggerTypeValue);
                    
                    if (effect == null) {
                        LOGGER.warning("无法添加动画效果到段落");
                    }
                } catch (Exception e) {
                    LOGGER.log(Level.SEVERE, "添加动画时发生错误", e);
                    // 文本已添加成功，即使动画添加失败也返回成功
                }
            }
            
            return new SetFormattedTextResult(true, "格式化文本设置成功");
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "设置格式化文本失败", e);
            return new SetFormattedTextResult(false, "设置格式化文本失败: " + e.getMessage());
        }
    }
    
    /**
     * 获取效果类型值
     */
    private static int getEffectTypeValue(String effectType) {
        if (effectType == null) {
            return EffectType.Fade;
        }
        return EFFECT_TYPE_MAP.getOrDefault(effectType.toUpperCase(), EffectType.Fade);
    }
    
    /**
     * 获取效果子类型值
     */
    private static int getEffectSubtypeValue(String effectSubtype) {
        if (effectSubtype == null) {
            return EffectSubtype.None;
        }
        return EFFECT_SUBTYPE_MAP.getOrDefault(effectSubtype.toUpperCase(), EffectSubtype.None);
    }
    
    /**
     * 获取触发类型值
     */
    private static int getTriggerTypeValue(String triggerType) {
        if (triggerType == null) {
            return EffectTriggerType.OnClick;
        }
        return TRIGGER_TYPE_MAP.getOrDefault(triggerType.toUpperCase(), EffectTriggerType.OnClick);
    }
}
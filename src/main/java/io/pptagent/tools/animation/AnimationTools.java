package io.pptagent.tools.animation;

import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.aspose.slides.EffectSubtype;
import com.aspose.slides.EffectTriggerType;
import com.aspose.slides.EffectType;
import com.aspose.slides.IAutoShape;
import com.aspose.slides.IEffect;
import com.aspose.slides.IParagraph;
import com.aspose.slides.IShape;
import com.aspose.slides.ISlide;
import com.aspose.slides.ITextFrame;
import com.aspose.slides.Presentation;
import io.pptagent.tools.PresentationManager;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;

/**
 * 动画相关工具函数
 */
public final class AnimationTools {
    private static final Logger LOGGER = Logger.getLogger(AnimationTools.class.getName());
    
    // 效果类型映射
    private static final Map<String, Integer> EFFECT_TYPE_MAP = Map.of(
        "APPEAR", EffectType.Appear,    // 出现
        "FADE", EffectType.Fade,        // 淡入淡出
        "FLOAT", EffectType.Float,      // 浮动
        "BOUNCE", EffectType.Bounce,    // 弹跳
        "WIPE", EffectType.Wipe,        // 擦除
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
        "AFTERPREVIOUS", EffectTriggerType.AfterPrevious,
        "WITHPREVIOUS", EffectTriggerType.WithPrevious
    );
    
    private AnimationTools() {
        // 私有构造函数防止实例化
    }
    
    /**
     * 表示动画参数的类
     */
    @Getter
    @Builder
    @AllArgsConstructor
    public static class AnimationParams {
        private final String effectType;      // 效果类型
        private final String effectSubtype;   // 效果子类型
        private final String triggerType;     // 触发类型
    }
    
    /**
     * 表示添加动画操作结果的类
     */
    @Getter
    @AllArgsConstructor
    public static class AddAnimationResult {
        private final boolean success;
        private final String message;
    }
    
    /**
     * 为形状添加动画效果
     * 
     * @param shapeIndex 形状索引
     * @param params 动画参数
     * @param slideIndex 幻灯片索引
     * @return 添加动画结果
     */
    public static AddAnimationResult addAnimation(int shapeIndex, AnimationParams params, int slideIndex) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                return new AddAnimationResult(false, "没有活动的演示文稿");
            }
            
            ISlide slide = pres.getSlides().get_Item(slideIndex);
            
            if (shapeIndex < 0 || shapeIndex >= slide.getShapes().size()) {
                return new AddAnimationResult(false, "无效的形状索引");
            }
            
            IShape shape = slide.getShapes().get_Item(shapeIndex);
            
            // 将字符串映射为枚举值
            int effectTypeValue = getEffectTypeValue(params.getEffectType());
            int effectSubtypeValue = getEffectSubtypeValue(params.getEffectSubtype());
            int triggerTypeValue = getTriggerTypeValue(params.getTriggerType());
            
            // 添加动画效果
            IEffect effect = slide.getTimeline().getMainSequence().addEffect(
                shape, effectTypeValue, effectSubtypeValue, triggerTypeValue);
                
            if (effect == null) {
                return new AddAnimationResult(false, "添加动画效果失败");
            }
            
            return new AddAnimationResult(true, "动画效果添加成功");
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "添加动画效果失败", e);
            return new AddAnimationResult(false, "添加动画效果失败: " + e.getMessage());
        }
    }
    
    /**
     * 为形状的指定段落添加动画效果
     * 
     * @param shapeIndex 形状索引
     * @param paragraphIndex 段落索引（从1开始）
     * @param params 动画参数
     * @param slideIndex 幻灯片索引
     * @return 添加动画结果
     */
    public static AddAnimationResult addParagraphAnimation(
            int shapeIndex, int paragraphIndex, AnimationParams params, int slideIndex) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                return new AddAnimationResult(false, "没有活动的演示文稿");
            }
            
            ISlide slide = pres.getSlides().get_Item(slideIndex);
            
            if (shapeIndex < 0 || shapeIndex >= slide.getShapes().size()) {
                return new AddAnimationResult(false, "无效的形状索引");
            }
            
            if (!(slide.getShapes().get_Item(shapeIndex) instanceof IAutoShape shape)) {
                return new AddAnimationResult(false, "指定的形状不是自动形状");
            }
            
            if (!shape.isTextBox()) {
                return new AddAnimationResult(false, "指定的形状没有文本框架");
            }
            
            ITextFrame textFrame = shape.getTextFrame();
            
            if (paragraphIndex < 1 || paragraphIndex > textFrame.getParagraphs().getCount()) {
                return new AddAnimationResult(false, "无效的段落索引");
            }
            
            IParagraph paragraph = textFrame.getParagraphs().get_Item(paragraphIndex);
            
            // 将字符串映射为枚举值
            int effectTypeValue = getEffectTypeValue(params.getEffectType());
            int effectSubtypeValue = getEffectSubtypeValue(params.getEffectSubtype());
            int triggerTypeValue = getTriggerTypeValue(params.getTriggerType());
            
            // 添加动画效果
            IEffect effect = slide.getTimeline().getMainSequence().addEffect(
                paragraph, effectTypeValue, effectSubtypeValue, triggerTypeValue);
                
            if (effect == null) {
                return new AddAnimationResult(false, "添加段落动画效果失败");
            }
            
            return new AddAnimationResult(true, "段落动画效果添加成功");
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "添加段落动画效果失败", e);
            return new AddAnimationResult(false, "添加段落动画效果失败: " + e.getMessage());
        }
    }
    
    /**
     * 获取效果类型值
     */
    private static int getEffectTypeValue(String effectType) {
        if (effectType == null) {
            return EffectType.Fade; // 默认为淡入淡出效果
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
            return EffectTriggerType.OnClick; // 默认为点击触发
        }
        return TRIGGER_TYPE_MAP.getOrDefault(triggerType.toUpperCase(), EffectTriggerType.OnClick);
    }
}
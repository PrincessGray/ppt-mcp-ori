package io.pptagent.tools.slides;

import com.aspose.slides.Presentation;
import com.aspose.slides.SlideLayoutType;
import io.pptagent.tools.PresentationManager;
import java.util.logging.Level;
import java.util.logging.Logger;

import lombok.Getter;
import lombok.AllArgsConstructor;

/**
 * 幻灯片相关工具函数
 */
public final class SlideTools {
    private static final Logger LOGGER = Logger.getLogger(SlideTools.class.getName());
    
    private SlideTools() {
        // 私有构造函数防止实例化
    }
    
    /**
     * 幻灯片布局类型枚举
     */
    @Getter
    public enum LayoutType {
        BLANK(SlideLayoutType.Blank),
        TITLE(SlideLayoutType.Title),
        TITLE_BODY(SlideLayoutType.TitleAndObject),
        TITLE_ONLY(SlideLayoutType.TitleOnly);
        
        private final byte value;
        
        LayoutType(byte value) {
            this.value = value;
        }

    }
    
    /**
     * 表示添加幻灯片结果的类
     */
    @Getter
    @AllArgsConstructor
    public static class AddSlideResult {
        private final boolean success;
        private final int slideIndex;
        private final String message;
    }
    
    /**
     * 从字符串转换为布局类型
     * 
     * @param layoutType 布局类型字符串
     * @return 布局类型
     */
    public static LayoutType layoutFromString(String layoutType) {
        if (layoutType == null) {
            return LayoutType.BLANK;
        }
        
        String type = layoutType.toUpperCase();
        return switch (type) {
            case "TITLE" -> LayoutType.TITLE;
            case "TITLEBODY" -> LayoutType.TITLE_BODY;
            case "TITLEONLY" -> LayoutType.TITLE_ONLY;
            default -> LayoutType.BLANK;
        };
    }
    
    /**
     * 添加新幻灯片
     * 
     * @param layoutType 布局类型(BLANK/TITLE/etc)
     * @return 新幻灯片索引，失败返回-1
     */
    public static int addSlide(String layoutType) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                return -1;
            }
            
            try {
                LayoutType layout = layoutFromString(layoutType);
                return pres.getSlides().addEmptySlide(
                    pres.getMasters().get_Item(0).getLayoutSlides().getByType(
                        layout.getValue()
                    )
                ).getSlideNumber() - 1;
            } catch (Exception e) {
                LOGGER.log(Level.SEVERE, "添加幻灯片失败", e);
                return -1;
            }
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "添加幻灯片失败", e);
            return -1;
        }
    }
    
    /**
     * 添加新幻灯片（增强版本）
     * 
     * @param layoutType 布局类型(BLANK/TITLE/etc)
     * @return 添加幻灯片结果
     */
    public static AddSlideResult addSlideEnhanced(String layoutType) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                return new AddSlideResult(false, -1, "没有活动的演示文稿");
            }
            
            try {
                LayoutType layout = layoutFromString(layoutType);
                int slideIndex = pres.getSlides().addEmptySlide(
                    pres.getMasters().get_Item(0).getLayoutSlides().getByType(
                        layout.getValue()
                    )
                ).getSlideNumber() - 1;
                return new AddSlideResult(true, slideIndex, "幻灯片添加成功");
            } catch (Exception e) {
                LOGGER.log(Level.SEVERE, "添加幻灯片失败", e);
                return new AddSlideResult(false, -1, "添加幻灯片失败: " + e.getMessage());
            }
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "添加幻灯片失败", e);
            return new AddSlideResult(false, -1, "添加幻灯片失败: " + e.getMessage());
        }
    }
    
    /**
     * 选择当前操作的幻灯片
     * 
     * @param slideIndex 幻灯片索引
     * @return 成功/失败状态
     */
    public static boolean selectSlide(int slideIndex) {
        return PresentationManager.getInstance().setCurrentSlideIndex(slideIndex);
    }
    
    /**
     * 表示选择幻灯片结果的类
     */
    @Getter
    @AllArgsConstructor
    public static class SelectSlideResult {
        private final boolean success;
        private final int slideIndex;
        private final String message;
    }
    
    /**
     * 选择当前操作的幻灯片（增强版本）
     * 
     * @param slideIndex 幻灯片索引
     * @return 选择幻灯片结果
     */
    public static SelectSlideResult selectSlideEnhanced(int slideIndex) {
        boolean success = PresentationManager.getInstance().setCurrentSlideIndex(slideIndex);
        if (success) {
            return new SelectSlideResult(true, slideIndex, "选择幻灯片成功");
        } else {
            return new SelectSlideResult(false, PresentationManager.getInstance().getCurrentSlideIndex(), 
                "选择幻灯片失败: 索引无效或没有活动的演示文稿");
        }
    }
}
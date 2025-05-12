package io.pptagent.tools;

import com.aspose.slides.Presentation;
import com.aspose.slides.SlideSizeType;
import com.aspose.slides.SlideSizeScaleType;
import lombok.Getter;

import java.util.Optional;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.logging.Level;
import java.util.logging.Logger;

/**
 * 演示文稿管理器 - 单例模式管理演示文稿实例
 * 使用Java 17特性重写
 */
public final class PresentationManager {
    private static final Logger LOGGER = Logger.getLogger(PresentationManager.class.getName());
    private static final PresentationManager INSTANCE = new PresentationManager();

    /**
     * -- GETTER --
     *  获取当前演示文稿实例
     *
     */
    @Getter
    private Presentation presentation;
    private final AtomicInteger currentSlideIndex = new AtomicInteger(0);
    
    private PresentationManager() {
        // 私有构造函数
    }
    
    /**
     * 获取PresentationManager单例实例
     * 
     * @return PresentationManager实例
     */
    public static PresentationManager getInstance() {
        return INSTANCE;
    }
    
    /**
     * 创建新的演示文稿
     * 
     * @return 操作结果
     */
    public boolean createPresentation() {
        try {
            if (presentation != null) {
                presentation.dispose();
            }
            presentation = new Presentation();
            
            // 设置演示文稿尺寸为16:9
            presentation.getSlideSize().setSize(1600, 900, SlideSizeScaleType.EnsureFit);
            
            currentSlideIndex.set(0);
            return true;
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "创建演示文稿失败", e);
            return false;
        }
    }

    /**
     * 获取当前演示文稿实例的Optional包装
     * 
     * @return 包含演示文稿的Optional
     */
    public Optional<Presentation> getPresentationOptional() {
        return Optional.ofNullable(presentation);
    }
    
    /**
     * 设置当前操作的幻灯片索引
     * 
     * @param index 幻灯片索引
     * @return 操作结果
     */
    public boolean setCurrentSlideIndex(int index) {
        return getPresentationOptional()
            .filter(pres -> index >= 0 && index < pres.getSlides().size())
            .map(pres -> {
                currentSlideIndex.set(index);
                return true;
            })
            .orElse(false);
    }
    
    /**
     * 获取当前操作的幻灯片索引
     * 
     * @return 当前幻灯片索引
     */
    public int getCurrentSlideIndex() {
        return currentSlideIndex.get();
    }
    
    /**
     * 释放资源
     */
    public void dispose() {
        if (presentation != null) {
            presentation.dispose();
            presentation = null;
        }
    }
}
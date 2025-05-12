package io.pptagent.tools.svg;

import java.util.logging.Level;
import java.util.logging.Logger;

import com.aspose.slides.IPPImage;
import com.aspose.slides.IPictureFrame;
import com.aspose.slides.ISlide;
import com.aspose.slides.ISvgImage;
import com.aspose.slides.ShapeType;
import com.aspose.slides.SvgImage;
import io.pptagent.tools.PresentationManager;

/**
 * SVG相关工具函数
 * 使用Java 17特性重写
 */
public final class SvgTools {
    
    private static final Logger LOGGER = Logger.getLogger(SvgTools.class.getName());
    
    private SvgTools() {
        // 私有构造函数防止实例化
    }
    
    /**
     * 表示SVG图像参数的记录
     */
    public record SvgParams(
        String svgContent,
        float x, 
        float y, 
        float width, 
        float height
    ) {}
    
    /**
     * 表示添加SVG图像结果的记录
     */
    public record AddSvgResult(boolean success, int imageIndex, String message) {}
    
    /**
     * 添加SVG图像
     * 
     * @param params SVG参数
     * @param slideIndex 幻灯片索引
     * @return 添加SVG图像结果
     */
    public static AddSvgResult addSvgImage(SvgParams params, int slideIndex) {
        try {
            return PresentationManager.getInstance().getPresentationOptional()
                .map(pres -> {
                    try {
                        // 验证SVG内容
                        if (params.svgContent() == null || params.svgContent().isEmpty()) {
                            return new AddSvgResult(false, -1, "SVG内容不能为空");
                        }
                        
                        ISlide slide = pres.getSlides().get_Item(slideIndex);
                        
                        // 创建SVG图像
                        ISvgImage svgImage = new SvgImage(params.svgContent());
                        
                        // 将SVG图像添加到演示文稿的图像集合中
                        IPPImage ppImage = pres.getImages().addImage(svgImage);
                        
                        // 创建图片框并添加SVG图像
                        IPictureFrame pictureFrame = slide.getShapes().addPictureFrame(
                            ShapeType.Rectangle, 
                            params.x(), 
                            params.y(), 
                            params.width(), 
                            params.height(), 
                            ppImage
                        );
                        
                        // 获取形状在幻灯片中的索引
                        int index = slide.getShapes().indexOf(pictureFrame);
                        return new AddSvgResult(true, index, "SVG图像添加成功");
                    } catch (Exception e) {
                        LOGGER.log(Level.SEVERE, "添加SVG图像失败", e);
                        return new AddSvgResult(false, -1, "SVG图像添加失败: " + e.getMessage());
                    }
                })
                .orElse(new AddSvgResult(false, -1, "没有活动的演示文稿"));
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "添加SVG图像失败", e);
            return new AddSvgResult(false, -1, "SVG图像添加失败: " + e.getMessage());
        }
    }
    
    /**
     * 根据SVG文件内容创建SVG图像并添加到指定幻灯片
     * 
     * @param svgContent SVG文件内容
     * @param x 图像X坐标
     * @param y 图像Y坐标
     * @param width 图像宽度
     * @param height 图像高度
     * @param slideIndex 幻灯片索引
     * @return 添加的图像索引，失败返回-1
     * @deprecated 使用{@link #addSvgImage(SvgParams, int)}替代
     */
    @Deprecated(since = "2.0", forRemoval = true)
    public static int addSvgImage(String svgContent, float x, float y, float width, float height, int slideIndex) {
        AddSvgResult result = addSvgImage(new SvgParams(svgContent, x, y, width, height), slideIndex);
        return result.success() ? result.imageIndex() : -1;
    }
}
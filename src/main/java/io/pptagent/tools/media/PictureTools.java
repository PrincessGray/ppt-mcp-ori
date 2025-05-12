package io.pptagent.tools.media;

import java.awt.Color;
import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.util.Base64;
import java.util.logging.Level;
import java.util.logging.Logger;
import com.aspose.slides.FillType;
import com.aspose.slides.IPPImage;
import com.aspose.slides.IPictureFrame;
import com.aspose.slides.ISlide;
import com.aspose.slides.Presentation;
import com.aspose.slides.ShapeType;
import io.pptagent.App;
import io.pptagent.tools.PresentationManager;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 图片相关工具函数
 */
public final class PictureTools {
    private static final Logger LOGGER = Logger.getLogger(PictureTools.class.getName());
    
    private PictureTools() {
        // 私有构造函数防止实例化
    }
    
    /**
     * 表示图片框参数的类
     */
    @Getter
    @AllArgsConstructor
    public static class PictureFrameParams {
        private final float x;
        private final float y;
        private final float width;
        private final float height;
        private final String borderColor;
        private final float borderWidth;
        private final float rotation;
        private final boolean lockAspectRatio;
    }
    
    /**
     * 表示添加图片框结果的类
     */
    @Getter
    @AllArgsConstructor
    public static class AddPictureFrameResult {
        private final boolean success;
        private final int frameIndex;
        private final String message;
    }
    
    /**
     * 获取完整的文件路径
     * 
     * @param filePath 文件路径
     * @return 完整的文件路径
     */
    private static String getFullPath(String filePath) {
        if (filePath == null || filePath.trim().isEmpty()) {
            throw new IllegalArgumentException("文件路径不能为空");
        }
        
        File file = new File(filePath);
        if (file.isAbsolute()) {
            return filePath; // 如果是绝对路径，直接返回
        }
        
        // 使用工作目录
        String workspace = App.getWorkspace();
        return new File(workspace, filePath).getAbsolutePath();
    }
    
    /**
     * 从文件路径添加保持原比例的图片框
     * 
     * @param imagePath 图片路径
     * @param x X坐标
     * @param y Y坐标
     * @param width 宽度
     * @param height 高度
     * @param slideIndex 幻灯片索引
     * @return 图片框索引，失败时返回-1
     */
    public static int addPictureFrameWithAspectRatio(String imagePath, float x, float y, float width, float height, int slideIndex) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                LOGGER.log(Level.SEVERE, "添加图片框失败: 没有活动的演示文稿");
                return -1;
            }
                
            ISlide slide = pres.getSlides().get_Item(slideIndex);
            String fullPath = getFullPath(imagePath);
            File imageFile = new File(fullPath);
                
            if (!imageFile.exists() || !imageFile.isFile()) {
                LOGGER.log(Level.SEVERE, "添加图片框失败: 图片文件不存在: " + fullPath);
                return -1;
            }
                
            IPPImage image = pres.getImages().addImage(new FileInputStream(imageFile));
                
            // 创建图片框，确保使用提供的宽度和高度
            IPictureFrame pictureFrame = slide.getShapes().addPictureFrame(
                ShapeType.Rectangle, 
                x, 
                y, 
                width, 
                height, 
                image
            );
                
            // 锁定纵横比以确保图片不会失真
            pictureFrame.getPictureFrameLock().setAspectRatioLocked(true);
                
            int frameIndex = slide.getShapes().indexOf(pictureFrame);
            return frameIndex;
        } catch (IOException e) {
            LOGGER.log(Level.SEVERE, "添加图片框失败: 文件读取错误", e);
            return -1;
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "添加图片框失败", e);
            return -1;
        }
    }
    
    /**
     * 从Base64字符串添加保持原比例的图片框
     * 
     * @param base64Image Base64编码的图片数据
     * @param x X坐标
     * @param y Y坐标
     * @param width 宽度
     * @param height 高度
     * @param slideIndex 幻灯片索引
     * @return 图片框索引，失败时返回-1
     */
    public static int addPictureFrameFromBase64(String base64Image, float x, float y, float width, float height, int slideIndex) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                LOGGER.log(Level.SEVERE, "添加图片框失败: 没有活动的演示文稿");
                return -1;
            }
                
            ISlide slide = pres.getSlides().get_Item(slideIndex);
            
            if (base64Image == null || base64Image.trim().isEmpty()) {
                LOGGER.log(Level.SEVERE, "添加图片框失败: Base64图片数据为空");
                return -1;
            }
            
            // 移除Base64前缀，如果有
            String imageData = base64Image;
            if (base64Image.contains(",")) {
                imageData = base64Image.substring(base64Image.indexOf(",") + 1);
            }
            
            // 解码Base64数据
            byte[] decodedBytes = Base64.getDecoder().decode(imageData);
            ByteArrayInputStream imageStream = new ByteArrayInputStream(decodedBytes);
            
            // 添加图片
            IPPImage image = pres.getImages().addImage(imageStream);
            
            // 创建图片框，确保使用提供的宽度和高度
            IPictureFrame pictureFrame = slide.getShapes().addPictureFrame(
                ShapeType.Rectangle, 
                x, 
                y, 
                width, 
                height, 
                image
            );
                
            // 锁定纵横比以确保图片不会失真
            pictureFrame.getPictureFrameLock().setAspectRatioLocked(true);
                
            int frameIndex = slide.getShapes().indexOf(pictureFrame);
            return frameIndex;
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "添加图片框失败: " + e.getMessage(), e);
            return -1;
        }
    }
    
    /**
     * 添加图片框并设置边框和旋转
     * 
     * @param imagePath 图片路径
     * @param x X坐标
     * @param y Y坐标
     * @param width 宽度
     * @param height 高度
     * @param borderColor 边框颜色（十六进制，如"#FF0000"）
     * @param borderWidth 边框宽度
     * @param rotation 旋转角度
     * @param slideIndex 幻灯片索引
     * @return 添加图片框结果
     */
    public static AddPictureFrameResult addPictureFrameWithFormat(String imagePath, float x, float y, 
                                                                 float width, float height, 
                                                                 String borderColor, float borderWidth,
                                                                 float rotation, int slideIndex) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                return new AddPictureFrameResult(false, -1, "没有活动的演示文稿");
            }
            
            try {
                ISlide slide = pres.getSlides().get_Item(slideIndex);
                String fullPath = getFullPath(imagePath);
                File imageFile = new File(fullPath);
                
                if (!imageFile.exists() || !imageFile.isFile()) {
                    return new AddPictureFrameResult(false, -1, "图片文件不存在: " + fullPath);
                }
                
                IPPImage image = pres.getImages().addImage(new FileInputStream(imageFile));
                
                // 创建图片框
                IPictureFrame pictureFrame = slide.getShapes().addPictureFrame(
                    ShapeType.Rectangle, 
                    x, y, 
                    width > 0 ? width : image.getWidth(), 
                    height > 0 ? height : image.getHeight(), 
                    image
                );
                
                // 设置边框
                if (borderColor != null && !borderColor.isEmpty()) {
                    pictureFrame.getLineFormat().getFillFormat().setFillType(FillType.Solid);
                    Color border = Color.decode(borderColor);
                    pictureFrame.getLineFormat().getFillFormat().getSolidFillColor().setColor(border);
                    pictureFrame.getLineFormat().setWidth(borderWidth);
                } else {
                    pictureFrame.getLineFormat().getFillFormat().setFillType(FillType.NoFill);
                }
                
                // 设置旋转角度
                if (rotation != 0) {
                    pictureFrame.setRotation(rotation);
                }
                
                // 锁定纵横比以确保图片不会失真
                pictureFrame.getPictureFrameLock().setAspectRatioLocked(true);
                
                int frameIndex = slide.getShapes().indexOf(pictureFrame);
                return new AddPictureFrameResult(true, frameIndex, "图片框添加成功");
            } catch (IOException e) {
                LOGGER.log(Level.SEVERE, "添加图片框失败: 文件读取错误", e);
                return new AddPictureFrameResult(false, -1, "添加图片框失败: 文件读取错误 - " + e.getMessage());
            } catch (Exception e) {
                LOGGER.log(Level.SEVERE, "添加图片框失败", e);
                return new AddPictureFrameResult(false, -1, "添加图片框失败: " + e.getMessage());
            }
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "添加图片框失败", e);
            return new AddPictureFrameResult(false, -1, "添加图片框失败: " + e.getMessage());
        }
    }
    
    /**
     * 从Base64字符串添加图片框并设置边框和旋转
     * 
     * @param base64Image Base64编码的图片数据
     * @param x X坐标
     * @param y Y坐标
     * @param width 宽度
     * @param height 高度
     * @param borderColor 边框颜色（十六进制，如"#FF0000"）
     * @param borderWidth 边框宽度
     * @param rotation 旋转角度
     * @param slideIndex 幻灯片索引
     * @return 添加图片框结果
     */
    public static AddPictureFrameResult addPictureFrameFromBase64WithFormat(String base64Image, float x, float y, 
                                                                           float width, float height, 
                                                                           String borderColor, float borderWidth,
                                                                           float rotation, int slideIndex) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                return new AddPictureFrameResult(false, -1, "没有活动的演示文稿");
            }
            
            try {
                ISlide slide = pres.getSlides().get_Item(slideIndex);
                
                if (base64Image == null || base64Image.trim().isEmpty()) {
                    return new AddPictureFrameResult(false, -1, "Base64图片数据为空");
                }
                
                // 移除Base64前缀，如果有
                String imageData = base64Image;
                if (base64Image.contains(",")) {
                    imageData = base64Image.substring(base64Image.indexOf(",") + 1);
                }
                
                // 解码Base64数据
                byte[] decodedBytes = Base64.getDecoder().decode(imageData);
                ByteArrayInputStream imageStream = new ByteArrayInputStream(decodedBytes);
                
                // 添加图片
                IPPImage image = pres.getImages().addImage(imageStream);
                
                // 创建图片框
                IPictureFrame pictureFrame = slide.getShapes().addPictureFrame(
                    ShapeType.Rectangle, 
                    x, y, 
                    width > 0 ? width : image.getWidth(), 
                    height > 0 ? height : image.getHeight(), 
                    image
                );
                
                // 设置边框
                if (borderColor != null && !borderColor.isEmpty()) {
                    pictureFrame.getLineFormat().getFillFormat().setFillType(FillType.Solid);
                    Color border = Color.decode(borderColor);
                    pictureFrame.getLineFormat().getFillFormat().getSolidFillColor().setColor(border);
                    pictureFrame.getLineFormat().setWidth(borderWidth);
                } else {
                    pictureFrame.getLineFormat().getFillFormat().setFillType(FillType.NoFill);
                }
                
                // 设置旋转角度
                if (rotation != 0) {
                    pictureFrame.setRotation(rotation);
                }
                
                // 锁定纵横比以确保图片不会失真
                pictureFrame.getPictureFrameLock().setAspectRatioLocked(true);
                
                int frameIndex = slide.getShapes().indexOf(pictureFrame);
                return new AddPictureFrameResult(true, frameIndex, "图片框添加成功");
            } catch (Exception e) {
                LOGGER.log(Level.SEVERE, "添加图片框失败", e);
                return new AddPictureFrameResult(false, -1, "添加图片框失败: " + e.getMessage());
            }
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "添加图片框失败", e);
            return new AddPictureFrameResult(false, -1, "添加图片框失败: " + e.getMessage());
        }
    }
}

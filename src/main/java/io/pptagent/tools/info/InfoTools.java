package io.pptagent.tools.info;

import java.util.ArrayList;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.aspose.slides.IAutoShape;
import com.aspose.slides.IGroupShape;
import com.aspose.slides.IParagraph;
import com.aspose.slides.IShape;
import com.aspose.slides.ISlide;
import com.aspose.slides.ITextFrame;
import com.aspose.slides.Presentation;

/**
 * 单文件测试类，包含InfoTools功能和测试代码
 */
public class InfoTools {
    
    private static final Logger LOGGER = Logger.getLogger(InfoTools.class.getName());
    
    /**
     * 表示形状信息的类
     */
    public static class ShapeInfo {
        private final int shapeIndex;          // 形状索引
        private final String shapeType;        // 形状类型
        private final float x;                 // X坐标
        private final float y;                 // Y坐标
        private final float width;             // 宽度
        private final float height;            // 高度
        private final boolean hasTextFrame;    // 是否有文本框架
        private final String textContent;      // 文本内容
        
        // 构造函数
        public ShapeInfo(int shapeIndex, String shapeType, float x, float y, 
                         float width, float height, boolean hasTextFrame, String textContent) {
            this.shapeIndex = shapeIndex;
            this.shapeType = shapeType;
            this.x = x;
            this.y = y;
            this.width = width;
            this.height = height;
            this.hasTextFrame = hasTextFrame;
            this.textContent = textContent;
        }
        
        // 手动添加getter方法
        public int getShapeIndex() { return shapeIndex; }
        public String getShapeType() { return shapeType; }
        public float getX() { return x; }
        public float getY() { return y; }
        public float getWidth() { return width; }
        public float getHeight() { return height; }
        public boolean isHasTextFrame() { return hasTextFrame; }
        public String getTextContent() { return textContent; }
    }
    
    /**
     * 获取幻灯片中所有形状的信息（包括群组中的形状）
     * 
     * @param pres Presentation对象
     * @param slideIndex 幻灯片索引
     * @return 形状信息列表
     */
    public static List<ShapeInfo> getShapesInfo(Presentation pres, int slideIndex) {
        List<ShapeInfo> shapesInfo = new ArrayList<>();
        try {
            if (pres == null) {
                LOGGER.severe("没有活动的演示文稿");
                return shapesInfo;
            }
            
            if (slideIndex < 0 || slideIndex >= pres.getSlides().size()) {
                LOGGER.severe("无效的幻灯片索引: " + slideIndex);
                return shapesInfo;
            }
            
            ISlide slide = pres.getSlides().get_Item(slideIndex);
            
            // 处理所有形状，包括群组中的形状
            processShapes(slide.getShapes(), shapesInfo);
            
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "获取形状信息失败", e);
        }
        return shapesInfo;
    }
    
    /**
     * 递归处理形状列表，包括群组中的嵌套形状
     */
    private static void processShapes(com.aspose.slides.IShapeCollection shapes, List<ShapeInfo> shapesInfo) {
        for (int i = 0; i < shapes.size(); i++) {
            try {
                IShape shape = shapes.get_Item(i);
                
                // 处理群组形状 - 递归处理其中的所有子形状
                if (shape instanceof IGroupShape) {
                    IGroupShape groupShape = (IGroupShape) shape;
                    processShapes(groupShape.getShapes(), shapesInfo);
                    continue; // 跳过将群组本身添加到结果列表
                }
                
                // 基本位置和大小信息
                float x = shape.getX();
                float y = shape.getY();
                float width = shape.getWidth();
                float height = shape.getHeight();
                
                // 获取形状类型
                String shapeType = getSimpleShapeType(shape);
                
                // 只处理自动形状中的文本
                boolean hasTextFrame = false;
                String textContent = "";
                
                if (shape instanceof IAutoShape) {
                    IAutoShape autoShape = (IAutoShape) shape;
                    if (autoShape.getTextFrame() != null) {
                        try {
                            String extractedText = extractTextFromTextFrame(autoShape.getTextFrame());
                            if (extractedText != null && !extractedText.trim().isEmpty()) {
                                hasTextFrame = true;
                                textContent = extractedText;
                            }
                        } catch (Exception e) {
                            LOGGER.log(Level.WARNING, "提取文本时出错: " + e.getMessage());
                        }
                    }
                }
                
                // 创建新的ShapeInfo对象并添加到列表
                ShapeInfo shapeInfo = new ShapeInfo(
                    shapesInfo.size(), // 使用动态索引
                    shapeType,
                    x,
                    y,
                    width,
                    height,
                    hasTextFrame,
                    textContent
                );
                
                shapesInfo.add(shapeInfo);
            } catch (Exception e) {
                LOGGER.log(Level.WARNING, "处理形状 #" + i + " 时出错: " + e.getMessage());
            }
        }
    }
    
    /**
     * 从文本框架提取文本
     */
    private static String extractTextFromTextFrame(ITextFrame textFrame) {
        if (textFrame == null) {
            return "";
        }
        
        StringBuilder text = new StringBuilder();
        
        try {
            // 直接使用getText方法获取整个文本框的内容
            text.append(textFrame.getText());
        } catch (Exception e) {
            // 如果直接获取文本失败，尝试逐段落提取
            try {
                int paragraphCount = textFrame.getParagraphs().getCount();
                
                for (int i = 1; i <= paragraphCount; i++) { // 注意: 从1开始，不是0
                    try {
                        IParagraph paragraph = textFrame.getParagraphs().get_Item(i);
                        if (paragraph != null && paragraph.getText() != null) {
                            text.append(paragraph.getText());
                            if (i < paragraphCount) {
                                text.append("\n");
                            }
                        }
                    } catch (Exception ex) {
                        // 跳过有问题的段落
                        LOGGER.log(Level.FINE, "跳过段落 #" + i);
                    }
                }
            } catch (Exception ex) {
                LOGGER.log(Level.FINE, "逐段落提取文本失败: " + ex.getMessage());
            }
        }
        
        return text.toString();
    }
    
    /**
     * 获取简化的形状类型
     */
    private static String getSimpleShapeType(IShape shape) {
        if (shape instanceof IAutoShape) {
            return "AUTO_SHAPE";
        } else if (shape instanceof com.aspose.slides.ITable) {
            return "TABLE";
        } else if (shape instanceof com.aspose.slides.IChart) {
            return "CHART";
        } else if (shape instanceof com.aspose.slides.IPictureFrame) {
            return "PICTURE";
        } else {
            return "OTHER";
        }
    }
    
    /**
     * 获取幻灯片数量
     */
    public static int getSlideCount(Presentation pres) {
        if (pres == null) {
            return 0;
        }
        return pres.getSlides().size();
    }   
}
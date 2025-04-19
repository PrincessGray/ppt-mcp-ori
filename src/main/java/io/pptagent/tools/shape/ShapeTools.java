package io.pptagent.tools.shape;

import java.awt.Color;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.aspose.slides.FillType;
import com.aspose.slides.IAutoShape;
import com.aspose.slides.IShape;
import com.aspose.slides.ISlide;
import com.aspose.slides.LineStyle;
import com.aspose.slides.ShapeType;
import com.aspose.slides.Presentation;
import io.pptagent.tools.PresentationManager;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 形状相关工具函数
 */
public final class ShapeTools {
    private static final Logger LOGGER = Logger.getLogger(ShapeTools.class.getName());
    
    private ShapeTools() {
        // 私有构造函数防止实例化
    }
    
    /**
     * 表示形状参数的类
     */
    @Getter
    @AllArgsConstructor
    public static class ShapeParams {
        private final float x;
        private final float y;
        private final float width;
        private final float height;
        private final String fillColor;
        private final String borderColor;
        private final float borderWidth;
    }
    
    /**
     * 表示添加形状结果的类
     */
    @Getter
    @AllArgsConstructor
    public static class AddShapeResult {
        private final boolean success;
        private final int shapeIndex;
        private final String message;
    }
    
    /**
     * 从字符串转换为形状类型值
     */
    public static int getShapeTypeValue(String type) {
        if (type == null) {
            return ShapeType.Rectangle;
        }
        
        String upperType = type.toUpperCase();
        if ("RECTANGLE".equals(upperType)) {
            return ShapeType.Rectangle;
        } else if ("ELLIPSE".equals(upperType) || "CIRCLE".equals(upperType)) {
            return ShapeType.Ellipse;
        } else if ("TRIANGLE".equals(upperType)) {
            return ShapeType.Triangle;
        } else if ("DIAMOND".equals(upperType)) {
            return ShapeType.Diamond;
        } else if ("STAR".equals(upperType)) {
            return ShapeType.FivePointedStar;
        } else if ("ARROW".equals(upperType)) {
            return ShapeType.RightArrow;
        } else {
            return ShapeType.Rectangle; // 默认为矩形
        }
    }
    
    /**
     * 添加基本形状
     * 
     * @param type 形状类型
     * @param params 形状参数
     * @param slideIndex 幻灯片索引
     * @return 添加形状结果
     */
    public static AddShapeResult addShape(String type, ShapeParams params, int slideIndex) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                return new AddShapeResult(false, -1, "没有活动的演示文稿");
            }
            
            try {
                ISlide slide = pres.getSlides().get_Item(slideIndex);
                
                // 确定形状类型
                int shapeTypeValue = getShapeTypeValue(type);
                
                // 创建形状
                IAutoShape shape = slide.getShapes().addAutoShape(
                    shapeTypeValue, 
                    params.getX(), 
                    params.getY(), 
                    params.getWidth(), 
                    params.getHeight()
                );
                
                // 设置填充颜色
                if (params.getFillColor() != null && !params.getFillColor().isEmpty()) {
                    shape.getFillFormat().setFillType(FillType.Solid);
                    Color fill = Color.decode(params.getFillColor());
                    shape.getFillFormat().getSolidFillColor().setColor(fill);
                } else {
                    shape.getFillFormat().setFillType(FillType.NoFill);
                }
                
                // 设置边框
                if (params.getBorderColor() != null && !params.getBorderColor().isEmpty()) {
                    shape.getLineFormat().getFillFormat().setFillType(FillType.Solid);
                    Color border = Color.decode(params.getBorderColor());
                    shape.getLineFormat().getFillFormat().getSolidFillColor().setColor(border);
                    shape.getLineFormat().setWidth(params.getBorderWidth());
                } else {
                    shape.getLineFormat().getFillFormat().setFillType(FillType.NoFill);
                }
                
                int shapeIndex = slide.getShapes().indexOf(shape);
                return new AddShapeResult(true, shapeIndex, "形状添加成功");
            } catch (Exception e) {
                LOGGER.log(Level.SEVERE, "添加形状失败", e);
                return new AddShapeResult(false, -1, "添加形状失败: " + e.getMessage());
            }
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "添加形状失败", e);
            return new AddShapeResult(false, -1, "添加形状失败: " + e.getMessage());
        }
    }
    
    /**
     * 表示线条参数的类
     */
    @Getter
    @AllArgsConstructor
    public static class LineParams {
        private final float x1;
        private final float y1;
        private final float x2;
        private final float y2;
        private final String color;
        private final float thickness;
    }
    
    /**
     * 表示添加线条结果的类
     */
    @Getter
    @AllArgsConstructor
    public static class AddLineResult {
        private final boolean success;
        private final int lineIndex;
        private final String message;
    }
    
    /**
     * 添加线条
     * 
     * @param params 线条参数
     * @param slideIndex 幻灯片索引
     * @return 添加线条结果
     */
    public static AddLineResult addLine(LineParams params, int slideIndex) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                return new AddLineResult(false, -1, "没有活动的演示文稿");
            }
            
            try {
                ISlide slide = pres.getSlides().get_Item(slideIndex);
                
                // 计算线条的位置和大小
                float x = Math.min(params.getX1(), params.getX2());
                float y = Math.min(params.getY1(), params.getY2());
                float width = Math.abs(params.getX2() - params.getX1());
                float height = Math.abs(params.getY2() - params.getY1());
                
                // 创建线条形状
                IShape line = slide.getShapes().addAutoShape(ShapeType.Line, x, y, width, height);
                
                // 设置线条颜色和粗细
                if (params.getColor() != null && !params.getColor().isEmpty()) {
                    line.getLineFormat().getFillFormat().setFillType(FillType.Solid);
                    Color lineColor = Color.decode(params.getColor());
                    line.getLineFormat().getFillFormat().getSolidFillColor().setColor(lineColor);
                    line.getLineFormat().setWidth(params.getThickness());
                    line.getLineFormat().setStyle(LineStyle.Single);
                }
                
                int lineIndex = slide.getShapes().indexOf(line);
                return new AddLineResult(true, lineIndex, "线条添加成功");
            } catch (Exception e) {
                LOGGER.log(Level.SEVERE, "添加线条失败", e);
                return new AddLineResult(false, -1, "添加线条失败: " + e.getMessage());
            }
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "添加线条失败", e);
            return new AddLineResult(false, -1, "添加线条失败: " + e.getMessage());
        }
    }
    
    /**
     * 添加基本形状（兼容旧版接口）
     * 
     * @param type 形状类型
     * @param x X坐标
     * @param y Y坐标
     * @param width 宽度
     * @param height 高度
     * @param fillColor 填充颜色
     * @param borderColor 边框颜色
     * @param borderWidth 边框宽度
     * @param slideIndex 幻灯片索引
     * @return 形状索引
     */
    public static int addShape(String type, float x, float y, float width, float height, 
                              String fillColor, String borderColor, float borderWidth, int slideIndex) {
        ShapeParams params = new ShapeParams(x, y, width, height, fillColor, borderColor, borderWidth);
        AddShapeResult result = addShape(type, params, slideIndex);
        return result.isSuccess() ? result.getShapeIndex() : -1;
    }
    
    /**
     * 添加线条（兼容旧版接口）
     * 
     * @param x1 起点X坐标
     * @param y1 起点Y坐标
     * @param x2 终点X坐标
     * @param y2 终点Y坐标
     * @param color 颜色
     * @param thickness 粗细
     * @param slideIndex 幻灯片索引
     * @return 形状索引
     */
    public static int addLine(float x1, float y1, float x2, float y2, String color, float thickness, int slideIndex) {
        LineParams params = new LineParams(x1, y1, x2, y2, color, thickness);
        AddLineResult result = addLine(params, slideIndex);
        return result.isSuccess() ? result.getLineIndex() : -1;
    }
}
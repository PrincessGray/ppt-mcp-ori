package io.pptagent.tools.background;

import java.awt.Color;

import com.aspose.slides.BackgroundType;
import com.aspose.slides.FillType;
import com.aspose.slides.IPPImage;
import com.aspose.slides.ISlide;
import com.aspose.slides.ISvgImage;
import com.aspose.slides.PictureFillMode;
import com.aspose.slides.Presentation;
import com.aspose.slides.SvgImage;
import io.pptagent.tools.PresentationManager;

/**
 * 背景相关工具函数
 */
public class BackgroundTools {
    
    /**
     * 设置纯色背景
     * 
     * @param color 颜色代码
     * @param slideIndex 幻灯片索引
     * @return 成功/失败状态
     */
    public static boolean setBackgroundColor(String color, int slideIndex) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                return false;
            }
            
            if (color == null || color.isEmpty()) {
                return false;
            }
            
            ISlide slide = pres.getSlides().get_Item(slideIndex);
            
            // 设置背景类型为自定义背景
            slide.getBackground().setType(BackgroundType.OwnBackground);
            
            // 设置填充类型为纯色
            slide.getBackground().getFillFormat().setFillType(FillType.Solid);
            
            // 设置背景颜色
            Color bgColor = Color.decode(color);
            slide.getBackground().getFillFormat().getSolidFillColor().setColor(bgColor);
            
            return true;
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }
    
    /**
     * 设置SVG背景
     * 
     * @param svgContent SVG内容
     * @param slideIndex 幻灯片索引
     * @return 成功/失败状态
     */
    public static boolean setBackgroundSvg(String svgContent, int slideIndex) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                return false;
            }
            
            if (svgContent == null || svgContent.isEmpty()) {
                return false;
            }
            
            ISlide slide = pres.getSlides().get_Item(slideIndex);
            
            // 创建SVG图像
            ISvgImage svgImage = new SvgImage(svgContent);
            
            // 将SVG图像添加到演示文稿的图像集合中
            IPPImage ppImage = pres.getImages().addImage(svgImage);
            
            // 设置背景类型为自定义背景
            slide.getBackground().setType(BackgroundType.OwnBackground);
            
            // 设置填充类型为图片
            slide.getBackground().getFillFormat().setFillType(FillType.Picture);
            
            // 设置图片填充模式为拉伸以适应幻灯片
            slide.getBackground().getFillFormat().getPictureFillFormat().setPictureFillMode(PictureFillMode.Stretch);
            
            // 设置背景图片
            slide.getBackground().getFillFormat().getPictureFillFormat().getPicture().setImage(ppImage);
            
            return true;
        } catch (Exception e) {
            e.printStackTrace();
            return false;
        }
    }
} 
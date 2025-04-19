package io.pptagent.tools.base;

import java.util.logging.Level;
import java.util.logging.Logger;
import java.io.File;

import com.aspose.slides.Presentation;
import com.aspose.slides.SaveFormat;
import io.pptagent.tools.PresentationManager;
import io.pptagent.App;

import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 基础工具函数
 */
public class BaseTools {
    private static final Logger LOGGER = Logger.getLogger(BaseTools.class.getName());
    
    private BaseTools() {
        // 私有构造函数防止实例化
    }
    
    /**
     * 创建新的空白演示文稿
     * 
     * @return 成功/失败状态
     */
    public static boolean createPresentation() {
        return PresentationManager.getInstance().createPresentation();
    }
    
    /**
     * 获取保存格式值
     * 
     * @param format 格式字符串
     * @return 保存格式值
     */
    private static int getFormatValue(String format) {
        if (format == null) {
            return SaveFormat.Pptx;
        }
        
        String upperFormat = format.toUpperCase();
        if ("PPT".equals(upperFormat)) {
            return SaveFormat.Ppt;
        } else if ("PDF".equals(upperFormat)) {
            return SaveFormat.Pdf;
        } else {
            return SaveFormat.Pptx; // PPTX为默认格式
        }
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
        File fullPath = new File(workspace, filePath);
        
        // 确保目标目录存在
        File parentDir = fullPath.getParentFile();
        if (parentDir != null && !parentDir.exists()) {
            parentDir.mkdirs();
        }
        
        return fullPath.getAbsolutePath();
    }
    
    /**
     * 保存演示文稿
     * 
     * @param filePath 保存路径
     * @param format 格式(PPTX/PPT/PDF)
     * @return 成功/失败状态
     */
    public static boolean savePresentation(String filePath, String format) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                return false;
            }
            
            int saveFormat = getFormatValue(format);
            pres.save(filePath, saveFormat);
            return true;
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "保存演示文稿失败: " + filePath, e);
            return false;
        }
    }
    
    /**
     * 表示保存结果的类
     */
    @Getter
    @AllArgsConstructor
    public static class SaveResult {
        private final boolean success;
        private final String message;
        private final String path;
    }
    
    /**
     * 保存演示文稿（增强版本）
     * 
     * @param filePath 保存路径
     * @param format 格式(PPTX/PPT/PDF)
     * @return 保存结果
     */
    public static SaveResult savePresentationEnhanced(String filePath, String format) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                return new SaveResult(false, "没有活动的演示文稿", filePath);
            }
            
            try {
                String fullPath = getFullPath(filePath);
                int saveFormat = getFormatValue(format);
                pres.save(fullPath, saveFormat);
                return new SaveResult(true, "保存成功", fullPath);
            } catch (Exception e) {
                LOGGER.log(Level.SEVERE, "保存演示文稿失败", e);
                return new SaveResult(false, "保存失败: " + e.getMessage(), filePath);
            }
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "保存演示文稿失败", e);
            return new SaveResult(false, "保存失败: " + e.getMessage(), filePath);
        }
    }
}
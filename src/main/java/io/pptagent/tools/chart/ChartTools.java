package io.pptagent.tools.chart;

import java.awt.Color;
import java.util.List;
import java.util.logging.Level;
import java.util.logging.Logger;

import com.aspose.slides.ChartType;
import com.aspose.slides.FillType;
import com.aspose.slides.IChart;
import com.aspose.slides.IChartDataPoint;
import com.aspose.slides.IChartDataWorkbook;
import com.aspose.slides.IChartSeries;
import com.aspose.slides.IDataLabel;
import com.aspose.slides.ISlide;
import com.aspose.slides.NullableBool;
import com.aspose.slides.Presentation;
import io.pptagent.tools.PresentationManager;

import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Getter;

/**
 * 图表相关工具函数
 */
public final class ChartTools {
    private static final Logger LOGGER = Logger.getLogger(ChartTools.class.getName());
    
    /**
     * 图表类型枚举
     */
    public enum ChartTypeEnum {
        COLUMN,     // 柱状图
        PIE,        // 饼图
        LINE        // 折线图
    }
    
    /**
     * 图表参数构建器类
     */
    @Getter
    @Builder
    public static class ChartParams {
        private final float x;                 // X坐标
        private final float y;                 // Y坐标
        private final float width;             // 宽度
        private final float height;            // 高度
        private final String title;            // 图表标题
        private final String backgroundColor;  // 背景颜色（十六进制颜色代码，如"#FFFFFF"）
        private final String borderColor;      // 边框颜色（十六进制颜色代码，如"#000000"）
        private final float borderWidth;       // 边框宽度
    }
    
    /**
     * 系列数据类
     */
    @Getter
    @Builder
    public static class SeriesData {
        private final String name;            // 系列名称
        private final List<Double> values;    // 系列值
        private final String color;           // 系列颜色（十六进制颜色代码，如"#FF0000"）
    }
    
    /**
     * 图表结果类
     */
    @Getter
    @AllArgsConstructor
    public static class ChartResult {
        private final boolean success;      // 是否成功
        private final int chartIndex;       // 图表索引
        private final String message;       // 结果消息
    }
    
    private ChartTools() {
        // 私有构造函数防止实例化
    }
    
    /**
     * 统一的图表创建入口
     * 
     * @param chartType 图表类型（COLUMN, PIE, LINE）
     * @param params 图表参数
     * @param categories 类别标签
     * @param seriesDataList 系列数据列表
     * @param slideIndex 幻灯片索引
     * @return 图表创建结果
     */
    public static ChartResult createChart(ChartTypeEnum chartType, ChartParams params, 
                                         List<String> categories, List<SeriesData> seriesDataList, 
                                         int slideIndex) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            if (pres == null) {
                return new ChartResult(false, -1, "没有活动的演示文稿");
            }
            
            switch (chartType) {
                case COLUMN:
                    return createColumnChart(params, categories, seriesDataList, slideIndex);
                case PIE:
                    // 饼图只使用第一个系列的数据
                    if (seriesDataList == null || seriesDataList.isEmpty()) {
                        return new ChartResult(false, -1, "饼图需要至少一个数据系列");
                    }
                    SeriesData pieData = seriesDataList.get(0);
                    List<String> sliceColors = null;
                    if (pieData.getColor() != null) {
                        // 对于饼图，如果提供了系列颜色，将其作为第一个切片的颜色
                        sliceColors = List.of(pieData.getColor());
                    }
                    return createPieChart(params, categories, pieData.getValues(), sliceColors, slideIndex);
                case LINE:
                    return createLineChart(params, categories, seriesDataList, slideIndex);
                default:
                    return new ChartResult(false, -1, "不支持的图表类型: " + chartType);
            }
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "创建图表失败", e);
            return new ChartResult(false, -1, "创建图表失败: " + e.getMessage());
        }
    }
    
    /**
     * 创建柱状图
     */
    private static ChartResult createColumnChart(ChartParams params, List<String> categories, 
                                              List<SeriesData> seriesDataList, int slideIndex) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            ISlide slide = pres.getSlides().get_Item(slideIndex);
            
            // 添加带有默认数据的柱状图
            IChart chart = slide.getShapes().addChart(ChartType.ClusteredColumn, 
                params.getX(), params.getY(), params.getWidth(), params.getHeight());
            
            // 设置图表标题和样式
            setupChartAppearance(chart, params);
            
            // 填充图表数据
            fillChartData(chart, categories, seriesDataList, ChartType.ClusteredColumn);
            
            // 返回图表索引
            int chartIndex = slide.getShapes().indexOf(chart);
            return new ChartResult(true, chartIndex, "柱状图创建成功");
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "创建柱状图失败", e);
            return new ChartResult(false, -1, "创建柱状图失败: " + e.getMessage());
        }
    }
    
    /**
     * 创建饼图
     */
    private static ChartResult createPieChart(ChartParams params, List<String> sliceLabels, 
                                           List<Double> values, List<String> sliceColors, int slideIndex) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            ISlide slide = pres.getSlides().get_Item(slideIndex);
            
            // 添加带有默认数据的饼图
            IChart chart = slide.getShapes().addChart(ChartType.Pie, 
                params.getX(), params.getY(), params.getWidth(), params.getHeight());
            
            // 设置图表标题和样式
            setupChartAppearance(chart, params);
            
            // 获取图表数据工作簿
            IChartDataWorkbook workbook = chart.getChartData().getChartDataWorkbook();
            
            // 清除默认系列和类别
            chart.getChartData().getSeries().clear();
            chart.getChartData().getCategories().clear();
            
            // 添加类别（饼图的切片）
            for (int i = 0; i < sliceLabels.size(); i++) {
                chart.getChartData().getCategories().add(
                    workbook.getCell(0, i + 1, 0, sliceLabels.get(i))
                );
            }
            
            // 添加系列（饼图只有一个系列）
            IChartSeries series = chart.getChartData().getSeries().add(
                workbook.getCell(0, 0, 1, "数据"), 
                chart.getType()
            );
            
            // 填充数据点
            for (int i = 0; i < values.size() && i < sliceLabels.size(); i++) {
                series.getDataPoints().addDataPointForPieSeries(
                    workbook.getCell(0, i + 1, 1, values.get(i))
                );
            }
            
            // 设置饼图颜色
            if (sliceColors == null || sliceColors.isEmpty()) {
                // 没有提供颜色，使用自动变色
                chart.getChartData().getSeriesGroups().get_Item(0).setColorVaried(true);
            } else {
                // 用户提供了颜色，使用指定颜色
                chart.getChartData().getSeriesGroups().get_Item(0).setColorVaried(false);
                
                // 设置数据点样式和标签
                for (int i = 0; i < Math.min(values.size(), sliceLabels.size()); i++) {
                    if (i < sliceColors.size() && sliceColors.get(i) != null) {
                        try {
                            IChartDataPoint point = series.getDataPoints().get_Item(i);
                            
                            // 设置标签显示百分比
                            IDataLabel label = point.getLabel();
                            label.getDataLabelFormat().setShowValue(true);
                            label.getDataLabelFormat().setShowPercentage(true);
                            
                            // 设置自定义颜色
                            Color sliceColor = Color.decode(sliceColors.get(i));
                            point.getFormat().getFill().setFillType(FillType.Solid);
                            point.getFormat().getFill().getSolidFillColor().setColor(sliceColor);
                        } catch (NumberFormatException e) {
                            LOGGER.log(Level.WARNING, "无效的颜色格式: " + sliceColors.get(i), e);
                        }
                    }
                }
            }
            
            // 显示引导线
            series.getLabels().getDefaultDataLabelFormat().setShowLeaderLines(true);
            
            // 返回图表索引
            int chartIndex = slide.getShapes().indexOf(chart);
            return new ChartResult(true, chartIndex, "饼图创建成功");
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "创建饼图失败", e);
            return new ChartResult(false, -1, "创建饼图失败: " + e.getMessage());
        }
    }
    
    /**
     * 创建折线图
     */
    private static ChartResult createLineChart(ChartParams params, List<String> categories, 
                                            List<SeriesData> seriesDataList, int slideIndex) {
        try {
            Presentation pres = PresentationManager.getInstance().getPresentation();
            ISlide slide = pres.getSlides().get_Item(slideIndex);
            
            // 添加带有默认数据的折线图
            IChart chart = slide.getShapes().addChart(ChartType.Line, 
                params.getX(), params.getY(), params.getWidth(), params.getHeight());
            
            // 设置图表标题和样式
            setupChartAppearance(chart, params);
            
            // 填充图表数据
            fillChartData(chart, categories, seriesDataList, ChartType.Line);
            
            // 返回图表索引
            int chartIndex = slide.getShapes().indexOf(chart);
            return new ChartResult(true, chartIndex, "折线图创建成功");
        } catch (Exception e) {
            LOGGER.log(Level.SEVERE, "创建折线图失败", e);
            return new ChartResult(false, -1, "创建折线图失败: " + e.getMessage());
        }
    }
    
    /**
     * 设置图表外观
     */
    private static void setupChartAppearance(IChart chart, ChartParams params) {
        // 设置图表标题
        if (params.getTitle() != null && !params.getTitle().isEmpty()) {
            chart.getChartTitle().addTextFrameForOverriding(params.getTitle());
            chart.getChartTitle().getTextFrameForOverriding().getTextFrameFormat().setCenterText(NullableBool.True);
            chart.getChartTitle().setHeight(20);
            chart.hasTitle();
        }
        
        // 设置背景颜色
        if (params.getBackgroundColor() != null && !params.getBackgroundColor().isEmpty()) {
            try {
                Color bgColor = Color.decode(params.getBackgroundColor());
                chart.getPlotArea().getFormat().getFill().setFillType(FillType.Solid);
                chart.getPlotArea().getFormat().getFill().getSolidFillColor().setColor(bgColor);
            } catch (NumberFormatException e) {
                LOGGER.log(Level.WARNING, "无效的背景颜色格式: " + params.getBackgroundColor(), e);
            }
        }
        
        // 设置边框
        if (params.getBorderColor() != null && !params.getBorderColor().isEmpty()) {
            try {
                Color borderColor = Color.decode(params.getBorderColor());
                chart.getPlotArea().getFormat().getLine().getFillFormat().setFillType(FillType.Solid);
                chart.getPlotArea().getFormat().getLine().getFillFormat().getSolidFillColor().setColor(borderColor);
                chart.getPlotArea().getFormat().getLine().setWidth(params.getBorderWidth());
            } catch (NumberFormatException e) {
                LOGGER.log(Level.WARNING, "无效的边框颜色格式: " + params.getBorderColor(), e);
            }
        }
    }
    
    /**
     * 填充图表数据（柱状图和折线图共用）
     */
    private static void fillChartData(IChart chart, List<String> categories, 
                                    List<SeriesData> seriesDataList, int chartType) {
        // 获取图表数据工作簿
        IChartDataWorkbook workbook = chart.getChartData().getChartDataWorkbook();
        
        // 清除默认系列和类别
        chart.getChartData().getSeries().clear();
        chart.getChartData().getCategories().clear();
        
        // 添加系列
        for (int i = 0; i < seriesDataList.size(); i++) {
            SeriesData seriesData = seriesDataList.get(i);
            chart.getChartData().getSeries().add(
                workbook.getCell(0, 0, i + 1, seriesData.getName()), 
                chartType
            );
        }
        
        // 添加类别
        for (int i = 0; i < categories.size(); i++) {
            chart.getChartData().getCategories().add(
                workbook.getCell(0, i + 1, 0, categories.get(i))
            );
        }
        
        // 填充数据
        for (int seriesIdx = 0; seriesIdx < seriesDataList.size(); seriesIdx++) {
            IChartSeries series = chart.getChartData().getSeries().get_Item(seriesIdx);
            SeriesData seriesData = seriesDataList.get(seriesIdx);
            List<Double> values = seriesData.getValues();
            
            for (int pointIdx = 0; pointIdx < values.size() && pointIdx < categories.size(); pointIdx++) {
                // 根据图表类型添加不同的数据点
                if (chartType == ChartType.ClusteredColumn) {
                    series.getDataPoints().addDataPointForBarSeries(
                        workbook.getCell(0, pointIdx + 1, seriesIdx + 1, values.get(pointIdx))
                    );
                } else if (chartType == ChartType.Line) {
                    series.getDataPoints().addDataPointForLineSeries(
                        workbook.getCell(0, pointIdx + 1, seriesIdx + 1, values.get(pointIdx))
                    );
                }
            }
            
            // 设置系列颜色
            if (seriesData.getColor() != null && !seriesData.getColor().isEmpty()) {
                try {
                    Color seriesColor = Color.decode(seriesData.getColor());
                    series.getFormat().getFill().setFillType(FillType.Solid);
                    series.getFormat().getFill().getSolidFillColor().setColor(seriesColor);
                } catch (NumberFormatException e) {
                    LOGGER.log(Level.WARNING, "无效的颜色格式: " + seriesData.getColor(), e);
                }
            }
            
            // 显示数据点值
            series.getLabels().getDefaultDataLabelFormat().setShowValue(true);
        }
    }
    
    /**
     * 添加柱状图到指定幻灯片
     * 
     * @param x X坐标
     * @param y Y坐标
     * @param width 宽度
     * @param height 高度
     * @param title 图表标题
     * @param categoryLabels 类别标签列表
     * @param seriesLabels 系列标签列表
     * @param seriesData 系列数据（二维数组，第一维是系列，第二维是该系列的值）
     * @param seriesColors 系列颜色（十六进制颜色代码，如"#FF0000"表示红色）
     * @param slideIndex 幻灯片索引
     * @return 图表对象的索引，失败返回-1
     */
    public static int addColumnChart(float x, float y, float width, float height, String title, 
                                   List<String> categoryLabels, List<String> seriesLabels, 
                                   List<List<Double>> seriesData, List<String> seriesColors, int slideIndex) {
        // 创建参数对象
        ChartParams params = ChartParams.builder()
            .x(x)
            .y(y)
            .width(width)
            .height(height)
            .title(title)
            .build();
        
        // 转换为SeriesData列表
        List<SeriesData> seriesDataList = convertToSeriesDataList(seriesLabels, seriesData, seriesColors);
        
        // 调用统一接口
        ChartResult result = createChart(ChartTypeEnum.COLUMN, params, categoryLabels, seriesDataList, slideIndex);
        return result.isSuccess() ? result.getChartIndex() : -1;
    }
    
    /**
     * 添加饼图到指定幻灯片
     * 
     * @param x X坐标
     * @param y Y坐标
     * @param width 宽度
     * @param height 高度
     * @param title 图表标题
     * @param sliceLabels 饼图切片标签
     * @param values 饼图数据值
     * @param sliceColors 饼图切片颜色（十六进制颜色代码，如"#FF0000"表示红色）
     * @param slideIndex 幻灯片索引
     * @return 图表对象的索引，失败返回-1
     */
    public static int addPieChart(float x, float y, float width, float height, String title, 
                                List<String> sliceLabels, List<Double> values, 
                                List<String> sliceColors, int slideIndex) {
        // 创建参数对象
        ChartParams params = ChartParams.builder()
            .x(x)
            .y(y)
            .width(width)
            .height(height)
            .title(title)
            .build();
        
        // 饼图只需要一个系列
        SeriesData seriesData = SeriesData.builder()
            .name("数据")
            .values(values)
            .build();
        
        // 调用统一接口
        ChartResult result = createChart(ChartTypeEnum.PIE, params, sliceLabels, List.of(seriesData), slideIndex);
        return result.isSuccess() ? result.getChartIndex() : -1;
    }
    
    /**
     * 添加折线图到指定幻灯片
     * 
     * @param x X坐标
     * @param y Y坐标
     * @param width 宽度
     * @param height 高度
     * @param title 图表标题
     * @param categoryLabels 类别标签列表
     * @param seriesLabels 系列标签列表
     * @param seriesData 系列数据（二维数组，第一维是系列，第二维是该系列的值）
     * @param seriesColors 系列颜色（十六进制颜色代码，如"#FF0000"表示红色）
     * @param slideIndex 幻灯片索引
     * @return 图表对象的索引，失败返回-1
     */
    public static int addLineChart(float x, float y, float width, float height, String title, 
                                 List<String> categoryLabels, List<String> seriesLabels, 
                                 List<List<Double>> seriesData, List<String> seriesColors, int slideIndex) {
        // 创建参数对象
        ChartParams params = ChartParams.builder()
            .x(x)
            .y(y)
            .width(width)
            .height(height)
            .title(title)
            .build();
        
        // 转换为SeriesData列表
        List<SeriesData> seriesDataList = convertToSeriesDataList(seriesLabels, seriesData, seriesColors);
        
        // 调用统一接口
        ChartResult result = createChart(ChartTypeEnum.LINE, params, categoryLabels, seriesDataList, slideIndex);
        return result.isSuccess() ? result.getChartIndex() : -1;
    }
    
    /**
     * 将原始数据转换为SeriesData列表
     */
    private static List<SeriesData> convertToSeriesDataList(List<String> seriesLabels, 
                                                           List<List<Double>> seriesData, 
                                                           List<String> seriesColors) {
        List<SeriesData> result = new java.util.ArrayList<>();
        
        for (int i = 0; i < seriesLabels.size() && i < seriesData.size(); i++) {
            String color = (seriesColors != null && i < seriesColors.size()) ? seriesColors.get(i) : null;
            
            SeriesData sd = SeriesData.builder()
                .name(seriesLabels.get(i))
                .values(seriesData.get(i))
                .color(color)
                .build();
                
            result.add(sd);
        }
        
        return result;
    }
}
package io.pptagent;

import com.aspose.slides.Picture;
import com.fasterxml.jackson.databind.ObjectMapper;
import io.modelcontextprotocol.server.McpAsyncServer;
import io.modelcontextprotocol.server.McpServer;
import io.modelcontextprotocol.server.McpServerFeatures;
import io.modelcontextprotocol.spec.McpSchema;
import io.modelcontextprotocol.server.transport.StdioServerTransportProvider;
import io.pptagent.mcp.*;
import io.pptagent.tools.PresentationManager;
import io.pptagent.tools.background.BackgroundTools;
import io.pptagent.tools.base.BaseTools;
import io.pptagent.tools.shape.ShapeTools;
import io.pptagent.tools.slides.SlideTools;
import io.pptagent.tools.svg.SvgTools;
import io.pptagent.tools.text.TextTools;
import io.pptagent.tools.text.TextTools.AddTextBoxResult;
import io.pptagent.tools.text.TextTools.SetFormattedTextResult;

import java.io.File;
import java.io.FileOutputStream;
import java.io.PrintStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.concurrent.CountDownLatch;

import lombok.Getter;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import reactor.core.publisher.Flux;
import reactor.core.publisher.Mono;

/**
 * PPT-Agent 应用程序入口
 */
public class App {
    private static final Logger log = LoggerFactory.getLogger(App.class);
    private static final String LOG_FILE_PATH = "pptagent.log";
    // 添加获取工作目录的方法
    @Getter
    private static String workspace; // 添加工作目录变量
    
    public static void main(String[] args) {
        // 获取WORKSPACE环境变量
        workspace = System.getenv("WORKSPACE");
        if (workspace == null || workspace.trim().isEmpty()) {
            workspace = System.getProperty("user.dir"); // 如果没有设置WORKSPACE，使用当前目录
            log.info("未设置WORKSPACE环境变量，使用当前目录: {}", workspace);
        } else {
            log.info("使用WORKSPACE环境变量: {}", workspace);
        }
        
        // 解析命令行参数
        if (args.length > 0 && args[0].equalsIgnoreCase("demo")) {
            // 重定向标准输出和标准错误到文件
            redirectSystemStreams();
            
            // 配置SLF4J日志级别（如果使用SLF4J Simple）
            System.setProperty("org.slf4j.simpleLogger.logFile", LOG_FILE_PATH);
            System.setProperty("org.slf4j.simpleLogger.defaultLogLevel", "INFO");
            
            // 执行演示模式
            runDemo();
        } else {
            // 启动MCP服务器模式
            startMcpServer();
        }
    }
    
    /**
     * 重定向System.out和System.err到文件
     */
    private static void redirectSystemStreams() {
        try {
            // 创建日志目录（如果不存在）
            File logFile = new File(LOG_FILE_PATH);
            if (!logFile.exists()) {
                logFile.getParentFile().mkdirs();
            }
            
            // 创建输出流
            FileOutputStream fileOutputStream = new FileOutputStream(logFile, true);
            PrintStream printStream = new PrintStream(fileOutputStream);
            
            // 备份原始输出流（在需要时可以恢复）
            PrintStream originalOut = System.out;
            PrintStream originalErr = System.err;
            
            // 重定向除System.in以外的流（保留标准输入用于JSON-RPC通信）
            System.setOut(new SpecialPrintStream(printStream, originalOut));
            System.setErr(printStream);
            
            // 记录重定向开始
            printStream.println("=== 日志开始 - " + new java.util.Date() + " ===");
        } catch (Exception e) {
            System.err.println("无法重定向输出流: " + e.getMessage());
            e.printStackTrace();
        }
    }
    
    /**
     * 特殊的PrintStream类，用于区分JSON-RPC消息和普通日志
     */
    private static class SpecialPrintStream extends PrintStream {
        private final PrintStream originalOut;
        
        public SpecialPrintStream(PrintStream logStream, PrintStream originalOut) {
            super(logStream);
            this.originalOut = originalOut;
        }
        
        @Override
        public void println(String x) {
            // 检查是否为JSON-RPC消息（通常以{开头）
            if (x != null && (x.startsWith("{") || x.startsWith("Content-Type:"))) {
                // JSON-RPC消息直接发送到原始输出流
                originalOut.println(x);
            } else {
                // 其他消息写入日志文件
                super.println(x);
            }
        }
        
        @Override
        public void print(String s) {
            // 同样区分JSON-RPC消息
            if (s != null && (s.startsWith("{") || s.startsWith("Content-Type:"))) {
                originalOut.print(s);
            } else {
                super.print(s);
            }
        }
    }
    
    /**
     * 启动MCP服务器
     */
    private static void startMcpServer() {
        log.info("正在启动PPT-Agent MCP服务器...");
        
        // 自动创建一个新的演示文稿
        boolean createResult = BaseTools.createPresentation();
        if (createResult) {
            log.info("已自动创建新的演示文稿");
            
            // 添加一个默认的空白幻灯片
            try {
                int slideIndex = SlideTools.addSlide("BLANK");
                log.info("已添加默认空白幻灯片，索引：{}", slideIndex);
            } catch (Exception e) {
                log.error("添加默认幻灯片失败: {}", e.getMessage());
            }
            
            // 启动服务器前注册一个关闭钩子，以便在服务器关闭时释放资源
            Runtime.getRuntime().addShutdownHook(new Thread(() -> {
                log.info("正在关闭PPT-Agent，释放资源...");
                PresentationManager.getInstance().dispose();
                log.info("资源已释放");
            }));
        } else {
            log.error("自动创建演示文稿失败");
        }
        
        // 创建传输提供者（使用STDIO）
        StdioServerTransportProvider transportProvider = new StdioServerTransportProvider(new ObjectMapper());
        
        // 创建异步服务器
        McpAsyncServer server = McpServer.async(transportProvider)
            .serverInfo("ppt-agent", "1.0.0")
            .capabilities(McpSchema.ServerCapabilities.builder()
                .tools(true) // 启用工具支持
                .build())
            .build();
            
        // 注册各种PPT操作工具
        registerAllTools(server)
            .doOnSuccess(v -> log.info("所有工具注册成功"))
            .doOnError(e -> log.error("工具注册失败: {}", e.getMessage()))
            .subscribe();
        
        log.info("PPT-Agent MCP服务器已启动，按Ctrl+C停止服务");
        
        // 阻塞主线程，保持服务器运行
        CountDownLatch latch = new CountDownLatch(1);
        try {
            latch.await();
        } catch (InterruptedException e) {
            log.info("服务器被中断: {}", e.getMessage());
        }
    }
    
    /**
     * 注册所有工具
     */
    private static Mono<Void> registerAllTools(McpAsyncServer server) {
        // 获取所有工具规范
        List<McpServerFeatures.AsyncToolSpecification> allTools = new ArrayList<>();
        
        // 添加演示文稿工具
        allTools.addAll(PresentationToolsRegistrar.createToolSpecifications());
        
        // 添加幻灯片工具
        allTools.addAll(SlideToolsRegistrar.createToolSpecifications());
        
        // 添加背景工具
        allTools.addAll(BackgroundToolsRegistrar.createToolSpecifications());
        
        // 添加形状工具
        allTools.addAll(ShapeToolsRegistrar.createToolSpecifications());
        
        // 添加文本工具 
        allTools.addAll(TextToolsRegistrar.createToolSpecifications());

        // 添加SVG工具
        allTools.addAll(SvgToolsRegistrar.createToolSpecifications());
        
        // 添加图表工具
        allTools.addAll(ChartToolsRegistrar.createToolSpecifications());

        // 添加动画工具
        allTools.addAll(AnimationToolsRegistrar.createToolSpecifications());

        // 添加图片工具
        allTools.addAll(PictureToolsRegistrar.createToolSpecifications());

        // 添加信息工具
        allTools.addAll(InfoToolsRegistrar.createToolSpecifications());
        
        // 逐个注册工具
        return Flux.fromIterable(allTools)
            .flatMap(toolSpec -> {
                McpSchema.Tool tool = toolSpec.tool();
                log.info("注册工具: {}", tool.name());
                return server.addTool(toolSpec);
            })
            .then();
    }
    
    /**
     * 运行PPT演示
     */
    private static void runDemo() {
        log.info("PPT-Agent 演示模式已启动");
        
        // 示例：创建一个简单的演示文稿
        log.info("创建演示文稿...");
        boolean result = BaseTools.createPresentation();
        log.info("创建演示文稿：{}", result ? "成功" : "失败");
        
        if (result) {
            try {
                // 添加标题幻灯片
                int titleSlideIndex = SlideTools.addSlide("TITLE");
                log.info("添加标题幻灯片，索引：{}", titleSlideIndex);
                
                // 设置标题幻灯片背景颜色
                BackgroundTools.setBackgroundColor("#F0F0F0", titleSlideIndex);
                
                // 添加标题文本框
                AddTextBoxResult titleBoxResult = TextTools.addTextBox(
                    50, 150, 600, 100, null, "#333333", 1, titleSlideIndex
                );
                int titleShapeIndex = titleBoxResult.getShapeIndex();
                log.info("添加标题文本框，索引：{}", titleShapeIndex);
                
                // 设置标题文本
                List<Map<String, Object>> titleText = new ArrayList<>();
                Map<String, Object> titlePart = new HashMap<>();
                titlePart.put("text", "PPT-Agent 演示");
                titlePart.put("fontName", "Arial");
                titlePart.put("fontSize", 40.0);
                titlePart.put("bold", true);
                titlePart.put("color", "#333333");
                titleText.add(titlePart);
                
                TextTools.setFormattedText(titleShapeIndex, titleText, titleSlideIndex);
                
                // 添加副标题文本框
                AddTextBoxResult subtitleBoxResult = TextTools.addTextBox(
                    50, 280, 600, 50, null, "#666666", 1, titleSlideIndex
                );
                int subtitleShapeIndex = subtitleBoxResult.getShapeIndex();
                log.info("添加副标题文本框，索引：{}", subtitleShapeIndex);
                
                // 设置副标题文本
                List<Map<String, Object>> subtitleText = new ArrayList<>();
                Map<String, Object> subtitlePart = new HashMap<>();
                subtitlePart.put("text", "使用Java实现的PowerPoint工具");
                subtitlePart.put("fontName", "Arial");
                subtitlePart.put("fontSize", 24.0);
                subtitlePart.put("color", "#666666");
                subtitleText.add(subtitlePart);
                
                TextTools.setFormattedText(subtitleShapeIndex, subtitleText, titleSlideIndex);
                
                // 添加内容幻灯片
                int contentSlideIndex = SlideTools.addSlide("BLANK");
                log.info("添加内容幻灯片，索引：{}", contentSlideIndex);
                
                // 设置内容幻灯片背景色
                BackgroundTools.setBackgroundColor("#FFFFFF", contentSlideIndex);
                
                // 添加标题文本框
                AddTextBoxResult contentTitleResult = TextTools.addTextBox(
                    50, 50, 600, 50, null, "#333333", 1, contentSlideIndex
                );
                int contentTitleIndex = contentTitleResult.getShapeIndex();
                
                // 设置内容标题文本
                List<Map<String, Object>> contentTitleText = new ArrayList<>();
                Map<String, Object> contentTitlePart = new HashMap<>();
                contentTitlePart.put("text", "多格式文本示例");
                contentTitlePart.put("fontName", "Arial");
                contentTitlePart.put("fontSize", 28.0);
                contentTitlePart.put("bold", true);
                contentTitlePart.put("color", "#333333");
                contentTitleText.add(contentTitlePart);
                
                TextTools.setFormattedText(contentTitleIndex, contentTitleText, contentSlideIndex);
                
                // 添加形状
                int shapeIndex = ShapeTools.addShape(
                    "RECTANGLE", 50, 120, 600, 150, 
                    "#E6F2FF", "#0066CC", 2, contentSlideIndex
                );
                log.info("添加矩形形状，索引：{}", shapeIndex);
                
                // 为形状设置多格式文本
                List<Map<String, Object>> formattedText = new ArrayList<>();
                
                Map<String, Object> part1 = new HashMap<>();
                part1.put("text", "这是粗体红色文本");
                part1.put("fontName", "Arial");
                part1.put("fontSize", 24.0);
                part1.put("bold", true);
                part1.put("color", "#FF0000");
                formattedText.add(part1);
                
                Map<String, Object> part2 = new HashMap<>();
                part2.put("text", "，这是斜体蓝色文本");
                part2.put("fontName", "Arial");
                part2.put("fontSize", 24.0);
                part2.put("italic", true);
                part2.put("color", "#0000FF");
                formattedText.add(part2);
                
                Map<String, Object> part3 = new HashMap<>();
                part3.put("text", "，这是普通黑色文本。");
                part3.put("fontName", "Arial");
                part3.put("fontSize", 24.0);
                part3.put("color", "#000000");
                formattedText.add(part3);
                
                SetFormattedTextResult textResult = TextTools.setFormattedText(shapeIndex, formattedText, contentSlideIndex);
                log.info("设置格式化文本: {}", textResult.isSuccess() ? "成功" : "失败");
                
                // 添加一个简单的SVG图像
                String svgContent = "<svg width='300' height='200'>" +
                                    "<rect width='100%' height='100%' fill='#f0f9ff'/>" +
                                    "<circle cx='150' cy='100' r='80' fill='#FF9900'/>" +
                                    "</svg>";
                int svgIndex = SvgTools.addSvgImage(svgContent, 200, 300, 300, 200, contentSlideIndex);
                log.info("添加SVG图像，索引：{}", svgIndex);
                
                // 保存演示文稿
                String outputPath = workspace + "/PPTAgent_Demo.pptx";
                boolean saved = BaseTools.savePresentation(outputPath, "PPTX");
                log.info("保存演示文稿到 {}: {}", outputPath, saved ? "成功" : "失败");
                
            } finally {
                // 释放资源
                PresentationManager.getInstance().dispose();
            }
        }
    }

}
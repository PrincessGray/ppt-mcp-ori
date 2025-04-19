package io.pptagent.mcp;

import com.fasterxml.jackson.databind.ObjectMapper;
import io.modelcontextprotocol.server.McpAsyncServer;
import io.modelcontextprotocol.server.McpServer;
import io.modelcontextprotocol.server.McpServerFeatures;
import io.modelcontextprotocol.spec.McpSchema;
import io.modelcontextprotocol.server.transport.StdioServerTransportProvider;

import java.util.ArrayList;
import java.util.List;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import reactor.core.publisher.Flux;
import reactor.core.publisher.Mono;

/**
 * PPT-Agent MCP服务器入口
 */
public class PptAgentServer {
    private static final Logger log = LoggerFactory.getLogger(PptAgentServer.class);
    
    public static void main(String[] args) {
        log.info("正在启动PPT-Agent MCP服务器...");
        
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
        
        log.info("PPT-Agent MCP服务器已启动");
        
        // 阻塞主线程
        try {
            Thread.currentThread().join();
        } catch (InterruptedException e) {
            log.error("服务器被中断: {}", e.getMessage());
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
        
        // 逐个注册工具
        return Flux.fromIterable(allTools)
            .flatMap(toolSpec -> {
                McpSchema.Tool tool = toolSpec.tool();
                log.info("注册工具: {}", tool.name());
                return server.addTool(toolSpec);
            })
            .then();
    }
}
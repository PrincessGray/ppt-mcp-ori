# PPT-Agent MCP服务器

PPT-Agent MCP服务器提供了通过大型语言模型（LLM）创建、编辑和管理PowerPoint演示文稿的能力。该服务器实现了Model Context Protocol (MCP)规范，使LLM能够通过标准接口与PPT-Agent工具进行交互。

## 功能概述

PPT-Agent MCP服务器提供以下功能：

- 创建和保存PowerPoint演示文稿
- 添加和管理幻灯片
- 设置幻灯片背景颜色和SVG背景
- 添加各种形状和线条
- 添加文本框和格式化文本
- 导入SVG图像
- 添加图片
- 获取幻灯片的已有信息


## 技术架构

- **核心框架**：使用Model Context Protocol SDK实现MCP服务器功能
- **传输层**：支持STDIO（标准输入/输出）传输，便于与Claude等LLM客户端集成
- **工具实现**：基于Aspose.Slides库提供的PowerPoint操作功能
- **异步处理**：使用Project Reactor实现异步响应处理

## 构建项目

使用Maven构建项目：

```bash
mvn clean package
```

这将生成两个JAR文件：
- 标准JAR文件：`target/pptagent-0.0.1-SNAPSHOT.jar`
- 包含所有依赖的可执行JAR文件：`target/pptagent-0.0.1-SNAPSHOT-jar-with-dependencies.jar`

## 运行服务器

### 使用Java命令运行

```bash
java -jar target/pptagent-0.0.1-SNAPSHOT-jar-with-dependencies.jar
```

### 与Claude Desktop集成

要在Claude Desktop中使用此MCP服务器，您需要编辑Claude Desktop的配置文件，将PPT-Agent添加为MCP服务器：

1. 找到Claude Desktop配置文件：
   - Windows: `%APPDATA%\Claude\claude_desktop_config.json`
   - macOS: `~/Library/Application Support/Claude/claude_desktop_config.json`

2. 添加PPT-Agent配置：

```json
{
  "mcpServers": {
    "ppt-agent": {
      "command": "java",
      "args": [
        "-jar",
        "/path/to/pptagent-0.0.1-SNAPSHOT-jar-with-dependencies.jar"
      ]
    }
  }
}
```

3. 重启Claude Desktop应用

明白！下面是按照你当前 README 的风格、只列出工具名称和一句话功能描述的补充版。你可以直接将这些内容插入到“可用工具”部分，丰富工具列表。

---

### 基础工具
- `createPresentation` - 创建新的空白演示文稿
- `savePresentation` - 保存演示文稿到指定路径

### 幻灯片工具
- `addSlide` - 添加新的幻灯片
- `selectSlide` - 选择当前操作的幻灯片

### 背景工具
- `setBackgroundColor` - 设置幻灯片背景颜色
- `setBackgroundSvg` - 设置幻灯片SVG背景

### 形状工具
- `addShape` - 添加形状到幻灯片
- `addLine` - 添加线条到幻灯片

### 图表工具
- `addColumnChart` - 添加柱状图到幻灯片
- `addPieChart` - 添加饼图到幻灯片
- `addLineChart` - 添加折线图到幻灯片

### 动画工具
- `addAnimation` - 为形状添加动画效果
- `addParagraphAnimation` - 为文本段落添加动画效果

### 图片/媒体工具
- `addPictureFrameWithAspectRatio` - 添加并保持原比例的图片
- `addPictureFrameFromBase64` - 通过Base64数据添加图片
- `addPictureFrameWithFormat` - 添加图片并设置边框和旋转
- `addPictureFrameFromBase64WithFormat` - 通过Base64数据添加图片并设置边框和旋转

### SVG工具
- `addSvgImage` - 添加SVG图像到幻灯片

### 文本工具
- `addTextBox` - 添加文本框到幻灯片
- `setFormattedText` - 设置形状的格式化文本

### 信息工具
- `getShapesInfo` - 获取幻灯片中所有形状的信息
- `getSlideCount` - 获取演示文稿的幻灯片数量


## 使用示例

以下是一些使用示例，展示如何通过LLM使用PPT-Agent工具：

1. 创建新的演示文稿：
   ```
   请创建一个新的演示文稿
   ```

2. 添加标题幻灯片：
   ```
   在演示文稿中添加一个标题幻灯片，然后添加标题"PPT-Agent演示"
   ```

3. 添加内容幻灯片：
   ```
   添加一个新的空白幻灯片，设置背景为浅蓝色，然后添加一个标题"主要功能"
   ```

## 技术详情

PPT-Agent MCP服务器使用以下技术：
- Java 16
- Aspose.Slides (用于PPT操作)
- Model Context Protocol SDK (用于MCP实现)
- Project Reactor (用于异步响应处理)

## 自定义和扩展

要添加新的工具，您需要：
1. 在现有工具类中添加新方法
2. 在对应的工具注册器中创建新的工具规范
3. 更新README中的工具列表

## 许可证

此项目使用 MIT 许可证。请参阅 LICENSE 文件获取详情。 
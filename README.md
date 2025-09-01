# ExcelBeautifier 📊✨

一个优雅的 Excel 和 CSV 文件美化与转换工具，让你的表格瞬间变得专业而美观！

![image-20250901140504687](https://s1.vika.cn/space/2025/09/01/b5c3a2bcd5fc4c9baca1e245a4f43069)

## 🌟 简介

ExcelBeautifier 是一款专为提升表格文件颜值而设计的工具，它能够自动将普通的 CSV 和 Excel 文件转换为格式精美、排版规范的 Excel 文档。无论是数据报表、统计分析还是日常办公，都能让你的表格焕然一新！

![image-20250901141218415](https://s1.vika.cn/space/2025/09/01/422c109b4dd74a33b9447959d8d1902d)

## 🚀 功能特点

- **✨ 一键美化**：自动为 Excel 文件添加专业样式，包括标题栏高亮、边框美化和对齐优化
- **🔄 格式转换**：轻松将 CSV 文件转换为美观的 Excel 格式
- **📏 智能调整**：自动计算并调整列宽，确保内容完美显示
- **🎨 专业配色**：采用商务风格的配色方案，让表格既美观又不失专业
- **🔢 批量处理**：支持同时处理多个文件，提高工作效率
- **💾 自动备份**：处理前自动创建备份文件，防止数据丢失
- **🖥️ 跨平台支持**：兼容 Windows、macOS 和 Linux 系统

## 📋 安装指南

### 前提条件

- Python 3.6 或更高版本
- pip 包管理器

### 安装步骤

1. 克隆本仓库

   ```bash
   git clone https://github.com/Bifish0/ExcelBeautifier.git
   cd ExcelBeautifier
   ```

2. 安装依赖库

   ```bash
   pip install -r requirements.txt
   ```

   （工具会自动检测并安装所需依赖：`openpyxl` 和 `colorama`）

## 📝 使用方法

1. 运行程序

   ```bash
   python ExcelBeautifier.py
   ```

2. 按照提示操作：

   - 输入需要处理文件的目录（默认当前目录）
   - 输入美化后文件的保存目录（默认源文件目录）
   - 从列表中选择要处理的文件（直接回车选择全部）

3. 等待处理完成，查看结果

## 🎨 美化效果展示

![image-20250901141332224](https://s1.vika.cn/space/2025/09/01/106c355486554e5c9080fe54667449a6)

## 🛠️ 技术细节

- 使用 `openpyxl` 库处理 Excel 文件
- 采用 `colorama` 实现跨平台的彩色终端输出
- 自动检测文件类型并应用相应处理逻辑
- 智能判断单元格内容类型以设置最佳对齐方式

## 🤝 贡献指南

1. Fork 本仓库
2. 创建特性分支 (`git checkout -b feature/amazing-feature`)
3. 提交更改 (`git commit -m 'Add some amazing feature'`)
4. 推送到分支 (`git push origin feature/amazing-feature`)
5. 打开 Pull Request

## 📄 许可证

本项目采用 MIT 许可证 - 详见 [LICENSE](https://www.doubao.com/chat/LICENSE) 文件

## 👨‍💻 关于作者

**一只鱼（Bifish）**

- GitHub: [@Bifishone](https://github.com/Bifishone)
- 专注于开发提升办公效率的小工具

------



💖 希望这款工具能为你的工作带来便利！如果喜欢，请给个星星支持一下吧！⭐

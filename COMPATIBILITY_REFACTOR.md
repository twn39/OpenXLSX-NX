# OpenXLSX 兼容性问题记录与重构指南

## 1. 已识别的兼容性红线 (Compatibility Redlines)
在现代 Excel (尤其是 macOS 版及 Office 365) 中，OpenXLSX 生成的文件触发“发现内容有问题”警告的核心原因如下：

### A. 物理 ZIP 结构：Data Descriptor 冲突
- **问题**：`minizip-ng` 默认在每个条目后写入 "Data Descriptor" (Extended Local Header)。
- **后果**：Excel 的 ZIP 解析器对小型 Office 文件的物理结构校验极其严格，Data Descriptor 的存在会导致解析器认为文件流不规范。
- **验证**：使用系统 `zip -X` 或旧版 `zippy` 生成的无描述符结构 (`extended local header: no`) 是 100% 兼容的。

### B. XML 声明头 (XML Declarations)
- **问题**：OpenXLSX 硬编码了 `<?xml version="1.0" ... standalone="yes"?>`。
- **后果**：Excel 倾向于接受**不带声明头**的 XML 条目（特别是 `[Content_Types].xml` 和 `_rels/.rels`）。带有 `standalone="yes"` 往往被视为格式异常。

### C. 自闭合标签空格 (Self-closing Tag Spacing)
- **问题**：PugiXML 默认生成 `<tag/>`。
- **后果**：Excel 遗留解析器偏好 `<tag />` (斜杠前带空格)。

### D. 模板污染 (Metadata Pollution)
- **问题**：内置模板包含修订 UID (`xr:uid`)、绝对路径 (`mc:AlternateContent`) 和空的 `sharedStrings.xml`。
- **后果**：这些元数据在文件被程序化修改后会产生逻辑矛盾，触发修复警告。

---

## 2. 增量功能补全蓝图 (Incremental Feature Roadmap)
回退到基于 `zippy` 的健康版本 (`ebabc87`) 后，需按顺序补全以下逻辑：

### 第一阶段：核心加固 (Core Hardening)
1. **XML 劫持清洗**：在 `XLZipArchive` 的写入接口处，强制剥离 `<?xml` 头部，并正则化 `/>` 前的空格。
2. **模板净化**：在 `XLDocument::create` 中加入清洗逻辑，移除 `xr` 命名空间和绝对路径。
3. **空文件回避**：保存时检测 `sharedStrings.xml` 是否包含有效条目，若无则不写入。

### 第二阶段：数据验证功能 (Data Validation) - *当前优先级*
1. **移植 `XLDataValidation` 类**：
   - 支持 `whole` (整数), `decimal` (小数), `list` (下拉列表), `date` (日期) 等类型。
   - 确保生成的 XML 属性按字母顺序排列 (如 `allowBlank`, `error`, `errorStyle`, `sqref`, `type`)。
2. **Worksheet 集成**：提供 `wks.addDataValidation(sqref)` 接口。

### 第三阶段：现代特性移植
1. **依赖升级**：逐步引入 `fast_float` (加速数值解析) 和 `GSL` (内存安全)，但需确保不破坏 ZIP 打包逻辑。
2. **样式增强**：支持更多现代对齐属性，同时保持 XML 结构的精简。

---

## 3. 验证准则
每个功能合并后，必须通过以下校验：
1. `zipinfo -v filename.xlsx | grep "extended local header"` 必须为 `no`。
2. `unzip -p filename.xlsx xl/workbook.xml` 头部严禁出现 `<?xml`。

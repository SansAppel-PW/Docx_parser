# 示例文档目录

这个目录包含了用于测试和演示 DOCX 解析器功能的示例文档。

## 目录说明

此目录专门存放示例文档，用于：
- 测试解析器功能
- 演示不同文档类型的解析效果  
- 提供使用参考

## 支持的文件类型

解析器支持处理以下类型的文档：
- **标准 DOCX 文档**: 基础的 Word 文档
- **包含表格的文档**: 带有复杂表格结构的文档
- **包含图片的文档**: 嵌入图像的文档
- **包含嵌入对象的文档**: 带有各种嵌入内容的文档
- **包含 SmartArt 图形的文档**: 带有智能图形的文档

## 使用方法

将您的测试文档放置在此目录中，然后使用以下方式测试解析器功能：

```python
from src.parsers.document_parser import DocumentParser

parser = DocumentParser()
result = parser.parse_document("Files/examples/your_document.docx")
```

## 注意事项

- 请确保示例文档不包含敏感信息
- 建议使用较小的文档进行测试
- 解析结果将保存在 `parsed_docs` 目录中

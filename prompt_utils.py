#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
AI提示工具模块：AI文档生成器
包含用于生成AI提示的各种函数
"""

def get_template_analysis_prompt(template_representation: str) -> str:
    """
    (This function is currently not used as stage 1 is deterministic).
    Generates a prompt for the AI to analyze the template structure.
    """
    return f"""
Please analyze the following structured Word template content and identify all the fields that need to be filled.
The structure is represented as a JSON object where keys are unique cell identifiers (e.g., 'table_0_row_1_col_2') and values are the text content of those cells.

Template Structure:
{template_representation}

Your task is to return a JSON object containing the unique identifiers for cells that are placeholders for data.
The value for each key should be an empty string.

Example Output:
{{
  "table_0_row_0_col_1": "",
  "table_0_row_1_col_1": ""
}}
"""

def get_fill_data_prompt(structured_template: str, input_data: str) -> str:
    """
    Generates a prompt for the AI to map input data to the template structure
    and create the final JSON for filling the document.
    """
    return f"""
你是一个专业的文档填写助手。你的任务是将JSON数据映射到结构化的Word文档模板中。

模板结构以JSON对象的形式提供，其中键是每个单元格的唯一标识符（如"table_0_row_1_col_2"），值是单元格的文本内容。
输入数据是一个单独的JSON对象。

你必须确定输入JSON中的哪些数据应该填入模板的哪些单元格中。
使用单元格的文本标签和位置（行、列）来解决歧义。例如，如果有多个"现场复核情况"单元格，使用左侧或上方单元格的标签来决定应该填入哪些内容。

**模板结构:**
```json
{structured_template}
```

**输入数据:**
```json
{input_data}
```

**重要任务要求:**
创建一个JSON对象，其中：
- 键是需要填充的模板结构中的唯一单元格标识符
- 值是来自输入数据的对应值
- **所有生成的内容必须使用中文，包括描述、说明和总结**

**高级映射指令:**
- **智能映射:** 不要只寻找完全匹配的关键词。运用你的知识和推理能力。例如，如果输入数据有"project_leader"字段，它应该映射到标记为"项目负责人"的模板单元格。
- **使用上下文:** 相邻单元格的文本至关重要。如果一个单元格是空的，但左侧单元格显示"项目名称"，你应该将项目名称映射到那个空单元格。
- **处理通用数据:** 如果输入数据包含通用字段如"description"、"notes"或"备注"，不要丢弃它。在模板中找到最合适的大型空文本区域来放置这些信息。运用你的最佳判断来确定这些内容的最合理位置。
- **摘要总结:** 如果输入数据中的值是复杂对象或数组，将其摘要为适合文档单元格的连贯中文字符串。

**图片处理:**
- 如果输入数据包含图片相关信息（如文件路径或描述），你应该：
  1. 在合适的单元格中创建中文描述性文本，如"详见附件：[图片描述]"。
  2. 在输出JSON的根级别创建一个特殊键 `__attachments__`。
  3. `__attachments__` 的值应该是一个对象数组，每个对象有 `title`（中文标题）和 `path`（文件路径）键。

**示例:**
如果模板有一个单元格 `table_0_row_4_col_1` 用于"损伤照片"，输入数据有 `{{"damage_photos_path": "/uploads/damage_pic.png"}}`，你的输出应该包含：
```json
{{
  "table_0_row_4_col_1": "详见附件：损伤现场照片",
  "__attachments__": [
    {{ "title": "附件：损伤现场照片", "path": "/uploads/damage_pic.png" }}
  ]
}}
```

**重要提醒:**
- 只包含需要用新数据填充的单元格的键值对
- 不要包含已经包含静态标签的单元格（如"项目名称:"）。你的工作是填充标签旁边或下方的单元格
- **所有输出内容必须使用中文**
- 只返回最终的JSON对象，不要包含任何解释或markdown格式
"""

def get_multimodal_extraction_prompt(fields_to_extract: str) -> str:
    """
    Generates a prompt for the AI to extract structured data from multimodal inputs.
    """
    
    return f"""
你是一个高度智能的AI助手，负责从各种来源（包括文本和图像）中提取结构化信息来填写表单。

你的目标是基于提供的内容填充一个JSON对象。仔细分析所有文本和图像。

**需要提取的信息架构:**
你的最终输出必须是一个JSON对象。以下是你需要提取的字段。尽量填写尽可能多的字段。如果找不到某个字段的信息，请在最终JSON中省略该字段。

```json
{fields_to_extract}
```

**操作指令:**
1. **分析所有输入:** 审查从文档中提取的文本和所有图像的内容。
2. **综合信息:** 将所有来源的信息组合起来获得最完整的图像。例如，图像可能显示文本中描述的损伤。
3. **处理图像:** 如果你发现相关图像，在输出JSON的适当字段中包含它们的文件路径。例如，如果你发现结构损伤的图像，你可以创建一个键如 `"damage_photos_path": ["/path/to/image1.png", "/path/to/image2.png"]`。文件路径会在文本内容中提供给你。
4. **输出JSON:** **只返回一个有效的JSON对象**，包含提取的数据。**所有提取的文本内容和描述必须使用中文**。不要包含任何其他文本、解释或markdown格式。

现在，分析以下内容并生成JSON对象。
""" 
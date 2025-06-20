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
You are an expert assistant for filling out Word documents. Your task is to map data from a JSON object to a structured Word document template.

The template's structure is provided as a JSON object where keys are unique identifiers for each cell (e.g., "table_0_row_1_col_2") and values are the cell's text content.
The input data is a separate JSON object.

You must determine which data from the input JSON should go into which cell of the template.
Use the cell's text label and its position (row, column) to resolve ambiguities. For example, if there are multiple "现场复核情况" cells, use the label in the cell to the left or above to decide which content goes where.

**Template Structure:**
```json
{structured_template}
```

**Input Data:**
```json
{input_data}
```

**Your Task:**
Create a single JSON object where:
- The keys are the unique cell identifiers from the template structure that should be filled.
- The values are the corresponding values from the input data.

**Few-shot Example:**
If the template has a cell `table_0_row_4_col_0` with text "原形制" and an adjacent cell `table_0_row_4_col_1` with text "现场复核情况", and the input data has a field `"original_condition_review": "The original structure was well-preserved."`, your output should include:
`"table_0_row_4_col_1": "The original structure was well-preserved."`

**IMPORTANT:**
- Only include key-value pairs for cells that need to be filled.
- Do not include cells that contain static labels.
- If a value in the input data is a complex object or array, summarize it into a coherent string suitable for a document cell.
- Return ONLY the final JSON object, without any explanations or markdown formatting.
""" 
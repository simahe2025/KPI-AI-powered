【GUI v2（指标自动匹配）使用说明】

一、更新点
- 自动匹配考核指标（1.1~5.1）并填入“对应考核指标”字段；
- 在 report.csv 记录“匹配指标”“匹配关键词”；
- 支持 .txt/.docx/.doc/.pdf；PDF扫描需先OCR；.doc需 textract/antiword。

二、使用步骤
1) pip install -r requirements_gui.txt
2) 运行 News2Template_GUI_v2_indicator.py
3) 拖拽或选择文件夹 → 开始处理 → 输出在 output_docs/

三、匹配原理（可自定义）
- 按关键词计分，取分最高的指标；
- 关键词可在脚本顶部 INDICATOR_RULES 中修改/扩展；
- 未命中则填“（待匹配）”。

四、注意
- 自动匹配仅作建议，最终以学院口径为准；
- 如需“优先级规则/黑白名单/多指标并列”，可继续扩展逻辑。

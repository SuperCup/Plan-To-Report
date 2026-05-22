# 规划转提报工具

这是一个本地 Windows 桌面工具，用于把业务人员设计的 Excel 转换模板执行成实际结果表。

当前版本是 MVP 骨架，重点支持：

- 选择转换模板
- 一次上传多个 Excel
- 解析所有 sheet，并按规则拆分合并单元格
- 给每个 sheet 指定类型，例如规划表、商品清单、预算分配表等
- 设置执行起始位置、结束位置和执行顺序
- 执行字段映射、固定值、拼接、正则解析、查表匹配、条件跳过
- 预览生成结果
- 导出 Excel
- 输出异常清单

## 本地运行

首次运行前安装依赖：

```powershell
python -m pip install -r requirements.txt
```

启动工具：

```powershell
python -m plan_to_report
```

如果未安装为包，也可以这样启动：

```powershell
$env:PYTHONPATH="src"
python -m plan_to_report
```

## 打包 Windows exe

```powershell
.\build.bat
```

打包产物在：

```text
dist\PlanToReport\PlanToReport.exe
```

打包后的可编辑模板目录在：

```text
dist\PlanToReport\templates
```

## 使用流程

1. 打开工具并选择转换模板。
2. 点击“批量上传 Excel”，一次选择多个 `.xlsx` 或 `.xlsm` 文件。
3. 工具会解析所有 sheet，并展示文件名、sheet 名、行列数和合并单元格提示。
4. 在 sheet 列表中选中一行，在 **Excel 内容预览** 区查看原始单元格；点击「设起始/设结束」后单击单元格，选取该 sheet 的执行区域（规划表通常从字段名所在行开始）。
5. 在 sheet 列表中指定 sheet 类型，例如 `规划表`、`商品清单`、`商品匹配逻辑`、`预算分配表`。
6. 设置执行顺序：
   - 逐行执行：首行为字段名，后续每行一条记录。
   - 逐列执行：首列为字段名，后续每列一条记录。
7. 点击“执行转换”，预览活动汇总表、活动对应UPC表和异常清单。
8. 点击“导出结果”生成 Excel。

## DeepSeek 配置（可选）

规划表转换可在向导 **步骤 ④** 启用 DeepSeek 辅助解析（机制拆分、歧义口味）。

配置保存在项目目录：

```text
config/app_settings.json
```

首次使用可复制示例文件：

```powershell
copy config\app_settings.example.json config\app_settings.json
```

在 `app_settings.json` 中填写：

- `api_key`：DeepSeek API Key
- `model`：默认 `deepseek-chat`
- `base_url`：默认 `https://api.deepseek.com`
- `enabled`：是否默认启用

也可在界面中修改后点击 **「保存 DeepSeek 配置」**；执行转换时会自动保存。若 Key 留空，将尝试读取环境变量 `DEEPSEEK_API_KEY`。

> `config/app_settings.json` 已加入 `.gitignore`，避免误提交密钥。

## 合并单元格拆分规则

- 单行横向合并：拆分后用原单元格内容填充每个单元格。
- 单列纵向合并：拆分后仅填充最后一个单元格。
- 多行多列合并：不自动猜测，进入合并单元格提示，由用户补充内容。

## 模板设计说明

模板放在 `templates` 目录，格式为 JSON。

一个模板包含：

- `inputs`：需要用户指定的 sheet 类型
- `outputs`：输出表定义
- `fields`：输出字段定义
- `filters`：行过滤规则

字段来源类型：

- `direct`：直接映射源字段
- `constant`：固定值
- `concat`：多个字段或固定值拼接
- `regex_extract`：正则提取
- `lookup`：从辅助表查表匹配
- `manual`：人工补充字段，并进入异常清单

建议先基于 `templates/规划转提报示例模板.json` 调整。

## 当前边界

当前尚未实现复杂能力：

- 横向勾选表自动转纵向
- 一行输入展开成多行输出
- 预算复杂分摊
- 可视化拖拽模板设计器

这些能力会在模板规则稳定后继续扩展。

# 生成《学生对教育者AI使用行为的社会认知：自我报告与源记忆的证据》PPT

本项目提供一个 Python 脚本，自动生成含占位图与图标（简化形状/箭头）的 PowerPoint 文件，内容覆盖 Introduction 的10页结构与演讲者备注，且预设统一配色和中文字体。

## 预览
- 16:9 比例
- 统一配色：蓝(#2563EB)、绿(#16A34A)、灰(#6B7280)
- 中文字体优先使用“Microsoft YaHei”，无法找到时回退 Arial
- 包含“角色生态图”（教师/学生/观察者三角关系 + 流向箭头）

## 使用方法
1. 安装依赖
   ```bash
   pip install python-pptx
   ```
2. 运行脚本
   ```bash
   python generate_intro_ppt.py
   ```
3. 输出文件
   - AIEd_Student_Social_Cognition_Intro_CN.pptx

## 自定义
- 若需替换主色/辅助色，可修改脚本顶部的 `COLORS`。
- 若需切换字体或字号，修改 `DEFAULT_FONT_NAME`、`TITLE_FONT_SIZE`、`BODY_FONT_SIZE`。
- 可在每页的 `NOTES_*` 常量中调整演讲者备注。

## 将产物提交到你的 GitHub
告诉我：
- 仓库：如 `archrivalexe/temp20250823`
- 分支：如 `main` 或目标分支
- 目标路径与文件名：如 `slides/AIEd_Student_Social_Cognition_Intro_CN.pptx`

我可以为你创建 PR 并把 PPT 放进去。

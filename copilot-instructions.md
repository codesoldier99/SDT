# Copilot Instructions For 本课程仓库

## Source Of Truth

- 唯一核心依据为 syllabus/2026-spring/传感器与检测技术授课计划_升级版 .pdf。
- 课程结构、学时、模块顺序、考核构成必须优先服从该 PDF。
- 不直接从历史课件、学生归档或旧代码复制内容进入正式产物。

## Content Style

- 主要语言为中文，关键术语保留英文，如 Allan Variance、ESKF、RTK、SO(3)。
- 面向工科学生，表述简洁、工程导向、避免空泛宣传。
- 每个教学材料都要交代工程目标、关键指标、实现路径和验证方式。

## Naming

- 讲义文件命名使用 lecture-xx-topic.md 或对应 PPT 名称。
- 实验文件命名使用 lab-xx-topic.md。
- 课程公告使用日期前缀，如 2026-05-16-site-initialization.md。
- 版本目录统一使用 2026-spring、2027-spring 这类格式。

## Required Templates

- 讲义必须包含：标题、学习目标、理论基础、工作原理、工程案例、总结、作业、讲者备注。
- 实验必须包含：目标、预习、设备、步骤、数据分析、报告要求、评分细则、安全提示、排障建议。
- 参考资料必须包含：难度分级、出版信息、摘要、适用讲次。
- 视频推荐必须包含：标题、链接、来源频道、对应讲次、推荐理由、自制需求。

## Migration Boundary

- 第一阶段仅迁入官方授课计划 PDF。
- 不迁入学生论文、成绩表、照片、往年归档、原始器件 PDF 和历史演示代码。
- 如需引用历史素材，只能在索引中说明来源目录，不直接复制为正式课程资产。

## Annual Update Workflow

1. 复制上一学年结构到新学年目录。
2. 根据新授课计划调整讲次、实验和考核。
3. 更新网站文档与课程公告。
4. 生成 Docusaurus 文档版本快照。
5. 记录新增技术主题，如 MEMS、IoT、AI-based detection、LiDAR。

## Quality Bar

- 所有新页面都应能映射到官方授课计划中的模块或产出要求。
- 站点内容必须通过 npm run build。
- 不写与课程无关的默认模板文本，不保留 Docusaurus 示例内容。
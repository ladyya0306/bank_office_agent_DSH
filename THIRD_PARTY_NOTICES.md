# 第三方组件与来源

- DeepSeek Harness：本机验收版本 `@deepseek-ai/dsh 0.1.5-rc.1`。代码依赖通过 npm 安装，准确解析结果见 `deepseek-harness-local/package-lock.json`，不把 node_modules 上传到仓库。
- 本地 `bank-approval-ui` 基于 DeepSeek 组件调整，保留其目录中的 MIT LICENSE 和版权说明。
- `toolV2/office_kit` 来自本项目旧办公工具并作增量修改，初始文件摘要保存在 `toolV2/engine-origin.json`。不表示该摘要等于现在的文件内容。
- Python 库的版本记录在 requirements 文件及验收环境记录中；各库遵循自身许可。可选 OCR、Office COM 等仍需要额外环境，不属于已验证填报主线。

仓库未代用户为原创项目内容选择新的开源许可证；发布源码不应被理解成替第三方变更许可。

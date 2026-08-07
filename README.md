# WPS 工程表清理助手
注：本助手完全由AI 生成，且表格内容过多会运行非常慢，视自己电脑配置而定

一个面向 **WPS 表格（Windows）** 的 JSA / XLAM 工程表辅助插件。

当前版本：**v1.8.0**

> 本项目主要用于工程造价、工程量清单、概预算、人材机汇总表等表格的批量整理与检查。

 别的表格应该也行，请自行测试/
 因为领导的神秘要求从而诞生的本插件（笑）

## 主要功能

- 删除完全空白行
- 按指定列为空删除行
- 清除匹配文字内容（不删除行列）
- 中文大项同名同号
- 数字序号顺位
- 清理短尾页表头
- 调整打印范围与分页线
- PDF 导出文字/打印风险检测
- 统一表格外框与内部线宽
- 批量删除工作表
- 全部工作表分页预览 / 普通视图

## 下载

普通用户建议直接从 **GitHub Releases** 下载最新版：

https://github.com/lyouroboros109/wps-engineering-table-cleaner/releases/latest

不要直接下载 `src` 目录来安装。

## 安装

1. 下载 `工程表清理助手_v1.8.0.xlam`。
2. 打开 WPS 表格。
3. 进入 **工具 → 加载项**。
4. 点击浏览并选择 `.xlam` 文件。
5. 勾选加载项并确认。
6. 重新打开 WPS 后，在顶部功能区寻找 **工程表清理**。

详细说明见：

- `docs/INSTALL.txt`

## 目录

```text
src/
  EngineeringTableCleaner.js   # JSA 主源码
  customUI.xml                 # WPS 功能区定义

release/
  工程表清理助手_v1.8.0.xlam

docs/
  INSTALL.txt
  CHANGELOG.md
  SELF_CHECK_v1.8.0.txt
  CHECKSUM_v1.8.0.txt

.github/
  ISSUE_TEMPLATE/
```

## 使用前提示

建议首次使用任何批量删除或结构调整功能时：

1. 先保存原始文件副本；
2. 在副本上测试；
3. 检查分页预览和打印预览；
4. 再处理正式文件。

不同 WPS 版本对 JSA、分页符、合并单元格边框等接口的行为可能存在差异。

## 反馈 Bug

请在仓库顶部进入 **Issues → New issue**。

建议附上：

- WPS 版本
- Windows 版本
- 插件版本
- 报错截图
- 操作步骤
- 脱敏后的测试表格（如可以提供）

## 源码

主源码位于：

```text
src/EngineeringTableCleaner.js
src/customUI.xml
```

## License

当前仓库建议在公开分享前选择一个开源许可证。

如果希望别人可以自由使用、修改和再次发布，通常可选择 **MIT License**。

如果希望别人修改并发布衍生版本时也必须继续开源，可考虑 **GPL-3.0**。

在确定许可证前，请阅读仓库中的 `LICENSE-OPTIONS.md`。

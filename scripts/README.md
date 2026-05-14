# xlsx → index.html 同步

## 文件


| 文件                                           | 作用                                                                                                                      |
| -------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------- |
| `sync_xlsx_to_index.py`                      | 读 `幻想少女公会招募图鉴.xlsx`，替换 `index.html` 里 `/* >>> SYNC_XLSX_DATA_BEGIN */` 与 `/* <<< SYNC_XLSX_DATA_END */` 之间的 `allData` 行 |
| `../.github/workflows/sync-recruit-data.yml` | 推送到 GitHub 后，若变更了 xlsx 或本脚本，自动同步并提交 `index.html`                                                                        |


## 表头（第一行）

须含且列名**完全一致**（顺序任意）：**位阶**、**角色**、**职业**、**种族**、**系别**、**地形**。

- **位阶**写入页面的 `稀有度`：`传说→3`、`史诗→2`、`稀有→1`、`普通→0`，或数字 `0~3`。
- **角色**写入页面的 `角色名`；**系别**→`属性`；**地形**→`地区`。

## 本地运行

在含 `index.html` 与 xlsx 的 `**tmp` 目录的上一级** 执行，或在任意目录指定路径：

```bash
pip install openpyxl
python scripts/sync_xlsx_to_index.py
# 或
python scripts/sync_xlsx_to_index.py --xlsx path/to/幻想少女公会招募图鉴.xlsx --index path/to/index.html
```

`--dry-run` 只统计行数，不写文件。


# Booklist System

一个完全在本地运行的书单管理网页。数据来自一个 Excel 底本，阅读状态、感想、日期和自定义清单保存在本机 JSON 文件中，不需要数据库或云端账号。

## 功能

- 书目浏览、搜索、分类与优先级筛选
- 想读 / 在读 / 读完 / 踢掉四种状态
- 读完后的评价、感想、开始日、读完日和阅读流水
- 批量修改与撤销、自定义清单、Markdown / JSON 导出
- 可选读取本机 Readest 阅读统计，只读、不上传
- 打开页面时可以自动同步 Excel 的最新内容

## 环境要求

- Python 3.10+
- `openpyxl`
- 任意现代浏览器

```bash
python -m pip install -r requirements.txt
```

## 快速开始

1. 初始化空书单底本：

```bash
python booklist.py init
```

2. 启动服务：

```bash
python server.py
```

3. 打开 <http://127.0.0.1:8770>。

也可以先查看示例底本：

```bash
cp sample/booklist.sample.xlsx booklist.xlsx
```

Windows 上可以直接双击 `start.vbs`，或运行：

```powershell
.\launch.ps1
```

## 维护书单

Excel 是书目底本，`booklist.py` 是维护入口：

```bash
python booklist.py list
python booklist.py stats
python booklist.py export
```

追加书目时传入 JSON 数组：

```bash
python booklist.py add books.json
```

`books.json` 示例：

```json
[
  {
    "书名": "The Design of Everyday Things",
    "作者": "Don Norman",
    "类别": "方法/工具",
    "优先级": "高",
    "为什么加": "理解可用性与反馈机制",
    "备注": "主题：设计"
  }
]
```

## 数据位置与隐私

- `booklist.xlsx`：书目底本
- `booklist.csv`：自动生成的只读镜像
- `data/我的标记.json`：网页中写入的个人阅读状态
- `data/covers/`：可选封面缓存

这些文件默认被 `.gitignore` 排除。应用只监听 `127.0.0.1`，不会把书单上传到任何服务器。Readest 集成只读取本机数据库；不需要时设置 `BOOKLIST_READEST=0`。

## 可选环境变量

| 变量 | 作用 |
| --- | --- |
| `BOOKLIST_PORT` | 服务端口，默认 `8770` |
| `BOOKLIST_XLSX` | 自定义 Excel 路径 |
| `BOOKLIST_CSV` | 自定义 CSV 镜像路径 |
| `BOOKLIST_DATA` | 自定义本地数据目录 |
| `BOOKLIST_READEST_DIR` | 自定义 Readest 数据目录 |
| `BOOKLIST_READEST=0` | 关闭 Readest 集成 |

## 许可证

代码以 MIT License 发布。图片素材来源见 [THIRD_PARTY_ASSETS.md](THIRD_PARTY_ASSETS.md)。

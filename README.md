# 层级架构分析工具

## 项目简介

该程序用于分析会员推荐关系，计算每个会员的**层级**、**下游人数**、**直接下游人数**及其**上游推荐路径**。支持从 `CSV`、`XLSX` 和 `XLS` 格式的文件读取会员信息，计算后会将结果保存为相同格式的文件。

### 功能特性

- 支持读取 `CSV`、`XLSX`、`XLS` 格式的文件。
- 计算每个会员的层级（即从根节点到该会员的推荐链长度）。
- 计算每个会员的总下游人数（包括直接和间接的所有下游）。
- 计算每个会员的直接下游人数（即该会员直接推荐的下游人数）。
- 提供每个会员的上游路径（即从根节点到该会员的推荐路径）。
- 输出文件会自动处理命名冲突，避免文件覆盖。

## 使用方法

### 环境依赖

在运行该程序前，请确保你已经安装了以下依赖库：

```
pip install pandas openpyxl xlrd xlwt
```

### 运行步骤

1. 运行程序：

    ```
   python Hierarchical_Analysis.py
    ```

3. 按照程序提示输入以下信息：

    - 输入文件路径（支持 `CSV`、`XLSX` 或 `XLS` 格式）。
    - 输出文件路径（如果留空，则默认保存至源文件路径，并自动添加 `_with_levels` 后缀）。
    - 会员 ID 列名。
    - 推荐人 ID 列名。

4. 程序将读取输入文件，进行会员层级、下游人数、直接下游人数和上游路径的计算，并保存结果到输出文件中。
    

### 示例

假设你有一个 `members.xlsx` 文件，文件中包含以下两列：

- `MemberID`：表示会员的 ID。
- `ReferrerID`：表示该会员的推荐人。

你可以使用如下命令来运行程序：

`python Hierarchical_Analysis.py`

根据提示输入：

- 输入文件路径：`members.xlsx`
- 输出文件路径：`members_with_levels.xlsx`
- 会员 ID 列名：`MemberID`
- 推荐人 ID 列名：`ReferrerID`

程序将计算每个会员的层级、下游人数、直接下游人数和上游路径，并保存到 `members_with_levels.xlsx` 文件中。

## 输出文件说明

输出文件将包含以下列：

- **Level**: 每个会员的层级，表示从根节点到该会员的推荐链长度。
- **Downstream_Count**: 每个会员的总下游人数，包含直接和间接推荐的所有下游会员。
- **Direct_Downstream_Count**: 每个会员的直接下游人数，表示该会员直接推荐的人数。
- **Upstream_Path**: 每个会员的上游路径，显示从根节点到该会员的推荐链。



# heuristic-constraint-optimizer / 启发式约束优化器

> 用 VBA 写成的**多目标、带约束的启发式优化器**：调整表格里每个物品的「数量」，在同时满足两个目标的同时，守住一组逐物品的硬约束。零依赖。
> A multi-objective, constraint-aware heuristic optimiser in VBA that adjusts each item's quantity to satisfy two goals at once while honouring hard per-item constraints. Zero dependencies.

---

## 痛点：它解决什么问题（中文）

很多排产、配单、预算分配问题的本质，都是同一个难题：**你要在「总量达标」和「某个单位属性指标达标」之间找平衡，同时还被一堆逐物品的硬规矩捆住手脚**。

具体说，它同时追求两个目标：

- **目标 A（聚合约束）**：所有物品的「总价值」尽量贴近一个目标总值；
- **目标 B（属性目标）**：所有物品的「单位聚合惩罚值」（成本、风险、或你自定义的任意单位分值）要压在某个上限之下。

而每个物品还被硬约束卡着：

- 单件价值必须落在 `[最小值, 最大值]` 区间；
- 数量必须是该物品「包装规格」的整数倍；
- 数量不能低于原数量的某个比例，也不能超过某个倍数；
- 你标记为「受保护」的物品，永远不动。

**手工做法的崩溃点**：这几个目标经常互相打架——为了压总量你可能把某物品数量调得过界，为了守区间又可能偏离总值。手工去试、反复拉扯，既慢又容易踩约束红线，还停不下来（死循环）。

**这个引擎的价值**：把一个通常只在重型优化库里才有的思路，做成**小巧、零依赖的 VBA 实现**：

1. **多目标搜索**：在「物理总量」和「可能与之冲突的属性目标」之间做平衡；
2. **边界内随机游走**：在安全阈值内用 `Rnd()` 试探，避免卡在糟糕的局部解；
3. **多级降级**：快速粗调 → 精准启发式循环 → 单物品公式微调 → 惩罚值优化 → 强制减少（即使目标冲突也保证给得出可用结果）；
4. **防死循环熔断**：墙钟时间上限 + 最大迭代次数 + 无进展即退出的三重保护，宏绝不会卡死。

## What problem it solves (English)

Many allocation / scheduling / budgeting problems share one hard shape: hit a target total value **and** keep a per-unit penalty (cost, risk, any score you define) under a ceiling, while each item is pinned by hard rules (unit-value bounds, package multiples, quantity-ratio limits, protected items). These goals often fight each other and manual tuning is slow and constraint-breaking. This engine packs a usually-heavy-library technique into a small, dependency-free VBA implementation: multi-objective search, bounded random walk, multi-stage fallback, and an anti-infinite-loop guard.

---

## 快速运行（中文）

**环境要求**

- Excel 或任何兼容的电子表格软件（支持 VBA）；
- 一张符合布局的表（物品行 + 数量列 + 单位价值列 + 惩罚值列等）。

**三步跑起来**

1. 打开 `examples/sample-data.xlsx`，按 `Alt + F11` 导入 `src/ConstrainedAllocator.bas`；
2. 运行 **`AdjustAllocation`** 做一遍快速、确定性的粗调；
3. 再运行 **`PreciseTune`** 做启发式精修。

> 示例数字是虚构的，无任何现实意义，只用来展示布局。换上你自己的数据、调好常量即可。

**套到自己的问题**：改 `ConstrainedAllocator.bas` 顶部的 `CONFIGURATION` 块——

- 把你的列映射到 `..._COL` 常量；
- 设置单位价值上下限、惩罚阈值、容差、循环熔断参数；
- 在 `IsProtectedItem` 里写上你自己的受保护物品关键词。

完整常量清单与第二个领域示例（预算分配）见 [`docs/Configuration.md`](docs/Configuration.md)。

## Quick start (English)

1. Open `examples/sample-data.xlsx`, `Alt+F11` → import `src/ConstrainedAllocator.bas`.
2. Run **`AdjustAllocation`** for a fast, deterministic pass.
3. Run **`PreciseTune`** for the heuristic refinement.

The sample numbers are fictional and only show the layout. Edit the `CONFIGURATION` block to map your columns, bounds, thresholds and loop guards; put your protected-item keywords in `IsProtectedItem`. See `docs/Configuration.md`.

---

## 工作原理 / How it works

引擎包含五块通用能力：多目标优化、边界随机试探（`Rnd()`）、多级降级引擎、防死循环熔断、硬约束集。整体流程是「快速粗调 → 精准微调」两段式：先快速把总量拉到目标附近，再用启发式循环在约束内精修，必要时走单物品公式微调和强制减少兜底，确保总能给出一个可用解。

## 配置要点 / Configuration

所有参数都在 `src/ConstrainedAllocator.bas` 顶部的 `CONFIGURATION` 块（列映射、单位价值上下限 `UNIT_VALUE_MIN/MAX`、惩罚上限 `PENALTY_CEILING`、容差 `TOL_COARSE/TOL_PRECISE`、时间上限 `TIME_LIMIT_SEC`、最大迭代 `MAX_ITER`、数量比例 `MIN_QTY_RATIO/MAX_QTY_RATIO`）。**参数名唯一真源是源码常量**，修改代码后请同步此处。完整说明与预算分配示例见 [`docs/Configuration.md`](docs/Configuration.md)。

## 适用场景 / Use cases

排产配单、预算分配、库存量优化——任何「在满足逐物品硬约束下，同时逼近总量目标与单位属性目标」的资源分配问题。

## 法律声明 / Legal & Non-affiliation

本仓库是**独立、通用的算法项目**，与任何商业产品、雇主或专有系统无隶属、无衍生关系。全部代码为通用优化模式的原始 clean-room 重写，不含任何第三方数据、品牌或业务逻辑。MIT 许可证仅覆盖作者在本仓库中的原创贡献；若你在专有系统中使用本代码，需自行确保符合所在单位政策。示例工作簿完全虚构，不含任何个人数据、客户记录或专有数据集。

## License

[MIT License](LICENSE) · Copyright (c) 2026 ivy92888-glitch.

---

<details><summary>维护说明（中英双语单文件，怎么避免漂移）</summary>

本文档是**单文件中英双语**：中文为权威详细描述，英文为简洁摘要（不强求逐句对等）。所有参数名（如 `UNIT_VALUE_MIN`、`PENALTY_CEILING`）直接取自源码常量，唯一真源是 `src/ConstrainedAllocator.bas` 中的 `CONFIGURATION` 块。修改代码逻辑时，只要同步更新源码常量名与本文档「配置要点」中的同名项即可，叙述性文字按需增删，不影响一致性。

</details>

# 全局复利与踩坑日志 (Global Compound & Pitfall Log)

> 本文件用于记录在 Workspace 中执行任务过程中犯过的错误、踩过的坑、走过的弯路，
> 以及经过验证后可复用的有效方法和最佳实践。
> **目的：让每一次教训都变成未来的复利，避免重复犯同样的错误。**

---

## 使用规则

1. **每次踩坑后**，及时记录到本文档
2. **每条记录包含：** 日期、分类、问题描述、原因分析、解决方案、教训总结
3. **分类标签：** 使用 `[技术]` `[流程]` `[沟通]` `[工具]` `[设计]` 等标签
4. **复用方法标记：** 在可复用的方法前加上 `✅ [复用]` 标记
5. **定期回顾：** 新项目开始前扫描本文档，留意相关分类的过往教训

---

## 条目格式模板

```
### [YYYY-MM-DD] [分类标签] 简短标题

**问题描述：**

**原因分析：**

**解决方案：**

**教训总结 / ✅ 复用方法：**
```

---

## 踩坑与教训记录

> （按时间倒序排列，最新在最上面）

### [2026-06-20] [技术] 内联事件处理器中的函数必须在全局作用域

**问题描述：**
在 "全部网格 MTD 排名" 表格中添加搜索栏时，`oninput="filterGridTable(this)"`
不生效，输入文字后无筛选反应。

**原因分析：**
`filterGridTable` 函数被定义在 `init()` 函数内部（作为嵌套函数），
而通过 `innerHTML` 创建的内联事件处理器只能调用全局作用域中的函数。
嵌套函数不在全局作用域内，因此事件触发时 `filterGridTable` 为 `undefined`。

**解决方案：**
将 `filterGridTable` 函数定义移动到 `function init()` 之前，确保它在全局作用域中。

**教训总结 / ✅ 复用方法：**
在 JavaScript 中使用 `innerHTML` 设置含 `onclick`、`oninput` 等内联事件属性时，
被调用的函数必须定义在全局作用域（`window` 对象）中。嵌套函数、模块私有函数
对内联事件处理器不可见。建议：如果使用 `innerHTML` 动态生成内容，优先考虑
全局函数定义，或使用 `addEventListener` 方式绑定事件。

---

### [2026-06-19] [技术] 错误地将代码插入到带 src 属性的 script 标签中

**问题描述：**
使用 `html.replace('</script>', new_code + '</script>', 1)` 将 JavaScript 代码
插入到 HTML 中时，代码被插入到了 CDN 脚本标签 `<script src="chart.js">` 的
内容区域中，导致所有内联 JavaScript 被浏览器忽略。

**原因分析：**
HTML 中 `<script src="...">` 带 `src` 属性的脚本标签，其标签体内的内联内容
会被浏览器忽略。`replace` 替换了第一个 `</script>`（CDN 脚本的闭标签），
而不是想替换的第二个 `</script>`（主脚本块闭标签），导致新代码被塞到了
CDN 标签体内。

**解决方案：**
使用 `html.rfind('</script>')` 找到最后一个（即主脚本块的）闭标签，
或通过更精确的模式匹配找到正确的插入点，而非替换第一个匹配。

**教训总结 / ✅ 复用方法：**
在 HTML 中动态插入 JavaScript 代码时：
1. 确认目标 `<script>` 块的位置（使用 `rfind` 找最后一个，而非 `find`）
2. 检查目标 `<script>` 标签是否带 `src` 属性（带 `src` 的标签忽略内联内容）
3. 更好的做法：在数据级别操作，使用 JSON 嵌入数据，而非拼接 JavaScript 文本

---

### [2026-06-19] [技术] Python f-string 中嵌套 JavaScript 时的花括号转义陷阱

**问题描述：**
在 Python 脚本中使用 `f'''...'''` 生成长 HTML/JavaScript 文本时，
JS 中的 `{` 和 `}`（如对象字面量、模板字符串、函数体）被 Python 误解析
为 f-string 的表达式分隔符，导致 `f-string: empty expression not allowed` 错误。

**原因分析：**
Python f-string 使用 `{expression}` 进行变量插值，而 JavaScript 频繁使用
`{` `}`（如代码块、对象、模板字符串）。f-string 中的 `{` 必须被写作 `{{`
才能输出字面量的 `{`，但自动转换所有 `{` 为 `{{` 很容易遗漏或过度。

**解决方案：**
避免在 Python f-string 中嵌入大规模 JavaScript 代码。改用以下模式：
- 使用普通 `'''...'''` 字符串（非 f-string）+ 后续 `.format()` 或 `replace()`
- 或者在 Python 中生成独立的 JSON 数据文件，在 HTML 中用 `<script>` 加载
- 对 HTML 的补丁操作使用 Python 的字符串替换（`replace`、正则）直接修改

**教训总结 / ✅ 复用方法：**
生成混合 Python/JavaScript 的 HTML 时，遵循"数据与表现分离"原则：
1. 用 Python 处理数据，输出 JSON
2. 在 HTML 中用 `<script>var DATA = {...}</script>` 嵌入 JSON
3. JS 处理 JSON 数据渲染 UI
4. 对已经生成好的 HTML 做增量修改时，用 Python 字符串操作直接编辑 HTML 文件，
   避免重新执行复杂的 f-string 代码生成流程

---

### [2026-06-19] [技术] JavaScript 闭包覆盖问题 — 多个实例共享同一全局函数

**问题描述：**
为 4月、5月、6月三个月份创建日期选择器（`bootDailyPicker`）时，
只有最后一个月（6月）的日期选择器正常工作，4月和5月的日期按钮
点击无响应。

**原因分析：**
每次调用 `bootDailyPicker(mk)` 都会覆盖全局 `window.selectDate` 函数。
由于 JavaScript 闭包的特性，最后一次调用定义的 `selectDate` 函数
引用的是最后一个月（6月）的局部变量（`MD`、`tableContainer`），
导致 4月、5月的日期按钮调用 `selectDate` 时使用的是 6 月的数据。

**解决方案：**
不用覆盖全局函数的方式，改用独立存储每个月份的渲染器：
```javascript
var dailyRenderers = {};
function bootDailyPicker(mk) {
  // ... 定义独立的渲染函数 ...
  dailyRenderers[mk] = { render: renderDailyTable, activeDate: dates[0] };
}
window.pickDate = function(mk2, date) {
  var r = dailyRenderers[mk2]; // 查找正确的月份
  if (!r) return;
  r.render(date);
};
```

**教训总结 / ✅ 复用方法：**
当需要为多个独立实例创建相似的 UI 组件时：
1. 不要使用共享的全局闭包函数
2. 使用 Map 或对象存储每个实例的独立渲染函数和状态
3. 用一个查找函数（`pickDate`）按实例标识符（`mk`）分派到正确的处理函数
4. 这样每个实例的闭包独立，互不干扰

---

### [2026-06-19] [技术] HTML 拼接时引号转义层级混淆

**问题描述：**
在 JavaScript 中通过 `innerHTML` 生成带事件处理器的 HTML 元素时，
`onclick="showGrid('网格名')"` 的单引号需要在多层转义中正确传递。

**原因分析：**
转义层级：HTML 属性需要用双引号或单引号 → HTML 属性值中的 JS 代码用引号
→ JS 字符串中的 JS 表达式用引号。三层嵌套容易出错。

**解决方案：**
简化事件处理器的传参方式：
```javascript
// 不好：onclick="filterGridTable('4月')" — 引号层级复杂
// 好：oninput="filterGridTable(this)" — 直接传 DOM 元素
// 更好：在 DOM 树中通过 parent/child 关系定位目标元素，避免传参
```

**教训总结 / ✅ 复用方法：**
动态生成含事件处理器的 HTML 时：
1. 优先传递 `this`（元素自身），让函数通过 DOM 遍历找到关联元素
2. 使用 `element.closest('.parent')`、`element.querySelector('.child')` 等
   代替传参
3. 避免在事件属性值中嵌入带引号的字符串参数

---

## 复用方法清单

> 以下为经过验证、可在未来项目中直接使用的方法。不含具体踩坑细节，仅做快速索引。

| 分类 | 方法 | 首次记录日期 | 适用场景 |
|------|------|-------------|----------|
| Python 数据处理 | openpyxl 读取 Excel 数据 | 2026-06-16 | 网格积分、统计报表等 Excel 数据分析场景 |
| Python HTML 生成 | 数据与表现分离：Python 处理数据 → JSON 嵌入 HTML → JS 渲染 | 2026-06-19 | 生成含大量数据的 HTML 看板/报告 |
| HTML 补丁 | 用 Python 字符串替换操作直接修改已生成的 HTML，避免 f-string 转义 | 2026-06-19 | 对复杂 HTML 做增量修改 |
| JS 多实例组件 | 使用 Map/对象存储每个实例的独立状态和渲染函数，避免闭包覆盖 | 2026-06-19 | 多实例 UI 组件（月度标签、多表格等） |
| JS 事件处理器 | `oninput="fn(this)"` + DOM 遍历，避免传参转义 | 2026-06-20 | 动态生成的 HTML 内联事件绑定 |
| JS 全局函数 | 通过 `innerHTML` 生成的内联事件处理器必须调用全局作用域的函数 | 2026-06-20 | `innerHTML` + `onclick/oninput` 事件绑定 |

---

## 版本记录

| 日期 | 版本 | 变更摘要 |
|------|------|----------|
| 2026-06-20 | v2.0 | 在 /Users/mr.g/Documents/Codex/Workspace/ 中重新创建，新增网格数据看板项目的5条踩坑记录 |

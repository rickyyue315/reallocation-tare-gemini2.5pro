UI 樣式規範（淺色系專業風格）

一、色彩方案
- 主背景：`#F5F5F5`–`#FFFFFF` 建議預設 `#F7F7F7`
- 品牌強調色：`#2B6CB0`（可依品牌調整）
- 輔助色（低飽和度）：
  - 專業藍：`#4B6FA9`
  - 靜謐青：`#3B7A78`
  - 典雅紫：`#6B6BB2`
- 文字與界面：
  - 主文字：`#333333`
  - 次文字：`#666666`
  - 失效/提示：`#999999`
  - 分隔線/邊框：`#E5E5E5`
  - 面板/卡片：`#FFFFFF`

二、版面配置
- 響應式網格：12 欄
  - 斷點建議：
    - `xs < 600px`：單欄堆疊，容器邊距 `16px`
    - `sm ≥ 600px`：12 欄，溝槽（gutter）`16px`
    - `md ≥ 960px`：12 欄，溝槽 `24px`
    - `lg ≥ 1280px`：最大容器寬度 `1200px`，溝槽 `24px`
    - `xl ≥ 1600px`：最大容器寬度 `1440px`，溝槽 `24px`
- 間距標準（8px 倍數）：`8, 16, 24, 32, 40, 48, 64, 96`
  - 內容區塊間距不小於 `24px`
  - 頁面邊距建議：桌面 `24–32px`，行動 `16–24px`
- 留白原則：確保主要內容周邊至少一層 `24px` 留白帶，避免元素擁擠

三、圖示系統（Material Design 風格）
- 尺寸：
  - 標準：`24x24px`
  - 重要操作：`32x32px`
- 顏色：
  - 一般功能：`#666666`
  - 重要功能/強調：品牌色 `#2B6CB0`
  - 失效/禁用：`#999999`
- 使用準則：
  - 圖示與文字間距 `8–12px`
  - 同一區塊內圖示風格一致（線性/填色）
  - 圖示僅作輔助，不替代關鍵文案

四、專業元素
- 陰影（elevation 2–4）：
  - `elev-2`：`0 2px 8px rgba(0,0,0,0.08)`
  - `elev-3`：`0 4px 12px rgba(0,0,0,0.10)`
  - `elev-4`：`0 8px 24px rgba(0,0,0,0.12)`
- 圓角：
  - 控制項/輸入框：`4px`
  - 卡片/按鈕：`8px`
- 按鈕層級：
  - 主要（Primary）：填色，背景品牌色，文字白色
  - 次要（Secondary）：描邊，邊框品牌色，文字品牌色，背景白
  - 第三（Tertiary）：文字按鈕，文字品牌色，無背景
  - 狀態：`hover` 加深 `8–12%`，`active` 降低亮度並加強陰影，`disabled` 文字 `#999999`、背景 `#EEEEEE`

五、字體與字級（建議）
- 字體：`Inter`（英數）、`Noto Sans TC`（繁中）
- 字級（px）：`12, 14, 16, 20, 24, 32`
- 行高：`1.4–1.6`，段落間距使用 8px 倍數

六、表單與表格
- 輸入框高度：`40–48px`，內距：水平 `12–16px`
- 表格行高：`48–56px`，欄間距：`24px`
- 空狀態：提供圖示與引導文案，留白 ≥ `24px`

七、可用性與可及性
- 對比：主要文案對背景達到 `AA`（4.5:1）以上
- 觸控目標最小尺寸：`40x40px`
- 鍵盤可達性：焦點清晰（外框 `2px` 品牌色）

八、CSS Token（範例，可納入設計系統）

```css
:root {
  /* Color */
  --color-bg: #F7F7F7;
  --color-surface: #FFFFFF;
  --color-border: #E5E5E5;
  --color-text: #333333;
  --color-text-muted: #666666;
  --color-text-disabled: #999999;
  --color-brand: #2B6CB0;
  --color-accent-blue: #4B6FA9;
  --color-accent-teal: #3B7A78;
  --color-accent-purple: #6B6BB2;

  /* Spacing */
  --space-8: 8px;
  --space-16: 16px;
  --space-24: 24px;
  --space-32: 32px;
  --space-40: 40px;
  --space-48: 48px;
  --space-64: 64px;
  --space-96: 96px;

  /* Radius */
  --radius-4: 4px;
  --radius-8: 8px;

  /* Elevation */
  --elev-2: 0 2px 8px rgba(0,0,0,0.08);
  --elev-3: 0 4px 12px rgba(0,0,0,0.10);
  --elev-4: 0 8px 24px rgba(0,0,0,0.12);
}

.btn-primary {
  background: var(--color-brand);
  color: #fff;
  border-radius: var(--radius-8);
  padding: 0 var(--space-16);
  height: 40px;
  box-shadow: var(--elev-3);
}

.btn-secondary {
  background: var(--color-surface);
  color: var(--color-brand);
  border: 1px solid var(--color-brand);
  border-radius: var(--radius-8);
  padding: 0 var(--space-16);
  height: 40px;
}
```

九、圖示使用指南
- 選擇 Material Icons，同類型場景採一致風格
- 文字與圖示對齊基線，保持 `8–12px` 間距
- 重要操作使用 `32x32px` 並以品牌色強調

十、範例頁面：見 `design/prototypes/` 目錄中的高保真原型圖（首頁、清單頁、詳情頁）
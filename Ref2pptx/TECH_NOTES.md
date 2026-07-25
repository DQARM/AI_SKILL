# TECH_NOTES — pptx(OOXML)技術備忘(ref2pptx 附錄)

<!-- version: 2026-07-26.01 -->
版本:`2026-07-26.01`(與 `SKILL.md` 的 `version` 同步——每次改動任一檔,兩檔版號一起更新成當天日期 + 當日流水號 `YYYY-MM-DD.NN`)。

本檔集中 `SKILL.md` 引用的 **pptx XML 層技術細節**。規則與流程在 `SKILL.md`;這裡回答「怎麼在 XML 層正確做到」。每一節都對應 SKILL.md 的一條規則,動手寫 XML / 程式產出 / 做程式化 QA 之前先讀對應小節。

---

## §1 在既有 pptx 新增投影片(加頁 / 拆頁,editing 路線)

用程式新增投影片,必須**同步更新四處**:

1. 新的 `ppt/slides/slideN.xml`,及其 `ppt/slides/_rels/slideN.xml.rels`(至少一條關係連到 `slideLayout`)。
2. `[Content_Types].xml` 加該 slide 的 `<Override PartName="/ppt/slides/slideN.xml" ContentType="...presentationml.slide+xml"/>`。
3. `ppt/_rels/presentation.xml.rels` 加一條 slide 關係。
4. `ppt/presentation.xml` 的 `<p:sldIdLst>` 加 `<p:sldId>`(用沒用過的 id),並排到正確位置——**頁面順序由 `sldIdLst` 決定,不是檔名**。

**最關鍵的雷:新關係的 `rId` 不可與 `presentation.xml.rels` 既有的撞號。**
既有 rId 常被 `slideMaster / notesMaster / presProps / viewProps / theme / tableStyles / customXml` 佔走(實例:某檔 `rId10–rId14` 已被 notesMaster、presProps、viewProps、theme、tableStyles 佔用)。**先掃出目前最大的 rId、從 +1 開始發號**。撞號會把 theme / notesMaster 等關係覆蓋掉,**整份簡報直接壞掉**——徵兆:投影片變成顯示母片佔位提示字(「按一下以編輯母片文字樣式」)、或跑出別頁的備忘稿內容。

**整份重產(如 `pptxgenjs` 一次建好所有頁)沒有撞號問題,首選重產**;只有要保留使用者手動修改時才走加頁。

---

## §2 normAutofit / fontScale(autofit 為什麼會騙人)

- **空的 `<a:normAutofit/>`(沒有 `fontScale` 屬性)在 PowerPoint 會以 100% 顯示、不縮**——文字用母片原始字級溢出框、蓋住下方元素。
- **LibreOffice 開檔會即時重算縮放**,所以用 LibreOffice 轉圖檢查會「看起來正常」,**掩蓋 PowerPoint 的爆框**。視覺 QA 不可只信轉圖。
- 要靠 autofit 縮,必須**嵌入實算的縮放**:
  `<a:normAutofit fontScale="80000" lnSpcReduction="10000"/>`
- **fontScale 計算**:`fontScale = floor(可容行數 ÷ 需要行數 × 100000)`,取整後再保守下修一點。
- **可容行數** ≈ 內文框可用高度 ÷(母片內文字級 × 1.2),再打 8 折留呼吸空間。例:框高 3.7 吋 ≈ 266pt、母片內文 18pt → 每行 ≈ 21.6pt → 約 12 行、打 8 折 ≈ 9–10 行。
- **需要行數** = 各段行數總和;二層縮排子項各算一行、長句折行也要估進去,不是只數頂層項。
- **程式化 QA**:解壓後掃每張 `ppt/slides/slideN.xml`,找 `<a:normAutofit` 且**沒有** `fontScale` 屬性者 → 阻擋(未通過)。

---

## §3 buAutoNum 自動編號清單

- **段落範本**(懸掛縮排 + 自動編號):
  `<a:pPr marL="457200" indent="-457200"><a:buFont typeface="+mj-lt"/><a:buAutoNum type="arabicPeriod"/></a:pPr>`
- **手打號碼 / 圈號的偵測與移除**:`<a:t>` 文字開頭符合以下者,刪掉前綴、只留正文:
  - 半形數字:`^\s*\d+[.、)]\s*`
  - 中文序號:`一、`、`(1)`、`(1)` 這類
  - **圈號(enclosed 數字)**:`①`(U+2460)~`⑳`(U+2473),及 `⑴⑵⑶`(U+2474 起)這類
- **跨頁接續連號**:同一清單跨多頁時,續頁用
  `<a:buAutoNum type="arabicPeriod" startAt="5"/>`
  (例:第一頁 1–4 → 續頁 `startAt="5"` 顯示 5–8),不要每頁從 1 重來。
- 無序要點用 `buChar`(項目符號),不用編號。
- **程式化 QA**:掃 `<a:t>` 開頭殘留上述號碼 / 圈號 → 阻擋;該用編號的清單(目錄、有序步驟)確認段落真的有 `buAutoNum`;跨頁清單確認續頁有 `startAt`。

---

## §4 notesSlide(備忘稿)

替某頁加備忘稿,需要:

1. 建 `ppt/notesSlides/notesSlideN.xml`,及其 `_rels/notesSlideN.xml.rels`(關係連到**該 slide** 與 **notesMaster**)。
2. 該 slide 的 `ppt/slides/_rels/slideN.xml.rels` 加一條連到 notesSlide 的關係。
3. `[Content_Types].xml` 加 notesSlide 的 `Override`。
4. **rId 一樣不可與既有撞號**(規則同 §1)。

備忘稿母片(`ppt/notesMasters/`)只管備忘稿列印外觀,**與投影片外觀 / 字級無關**——投影片一律看 `ppt/slideMasters/` + `ppt/slideLayouts/`。

---

## §5 版面安全程式化掃描(出界 / 重疊)

- **座標換算**:讀每張 `ppt/slides/slideN.xml` 各形狀的 `<a:off x= y=>`(左上角)與 `<a:ext cx= cy=>`(寬高),單位 EMU,**÷ 914400 = 吋**。
- **出界判定**:形狀任一邊落在畫面外(x 或 y < 0,或 x+cx / y+cy 超過投影片寬 / 高;16:9 = 13.333 × 7.5 吋,其他比例照換算)即為錯——移回畫面內或刪除。實例:曾有進度條標籤被算到 y ≈ 183 吋,一份簡報殘留 238 個出界形狀,把編輯檢視的捲動範圍撐長,滑鼠滾輪滾不到下一頁。
- **重疊判定**(座標層):標題下緣 ≤ 內文框上緣;三方視角面板上緣 > 內文框下緣(且留實際間距);進度條不與頁碼重疊。
- **爆框判定**(文字層):座標不重疊不代表沒蓋到——內文「實際文字」可能溢出框(見 §2)。用「需要行數 vs 可容行數」複核,並配合 §2 的空 normAutofit 掃描。
- 渲染確認可用 LibreOffice 轉圖,但**只能當輔助**(§2 的陷阱)。

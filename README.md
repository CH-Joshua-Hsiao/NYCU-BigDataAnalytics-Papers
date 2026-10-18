# 巨量資料分析：論文報告

本專案為 NYCU 巨量資料分析課程的論文報告資料，介紹兩篇 2026 年推薦系統領域的論文，並提供自動產生投影片的 Python 腳本與約 15 分鐘的中文講稿。

## 報告論文

| 代號 | 論文 | 會議 | 程式碼 |
|---|---|---|---|
| Paper A | LISRec: Modeling User Preferences with Learned Item Shortcuts for Sequential Recommendation（`3770854.3780337.pdf`） | KDD '26 | [NEUIR/LISRec](https://github.com/NEUIR/LISRec) |
| Paper B | Modeling Stage-wise Evolution of User Interests for News Recommendation（`2603.10471v1.pdf`） | WWW '26 | [yikeJin/SEIN](https://github.com/yikeJin/SEIN) |

兩篇論文都在探討同一個問題：使用者的互動紀錄能多準確地反映真正的喜好？

- **Paper A（語意）**：歷史紀錄中有隨手點擊的雜訊。以物品文字語意建立個人化相似度圖，只保留最大連通分量作為「捷徑」來建模使用者。
- **Paper B（時間）**：使用者興趣會隨時間改變。以全域圖捕捉長期偏好，再把歷史切成時段子圖，用 LSTM 與 self-attention 建模興趣演變。

## 檔案說明

| 檔案 | 說明 |
|---|---|
| `make_slides.py` | 從兩篇 PDF 裁切圖表，並產生 16:9 投影片 |
| `講稿.md` | 約 15 分鐘的中文講稿，每張投影片一段，附建議時間 |
| `paper_presentation.pptx` | 產生好的投影片（13 張，備忘稿已放入講稿） |
| `LISRec_詳解.pdf` | Paper A（LISRec）的中文詳解 |
| `StageWise新聞推薦_詳解.pdf` | Paper B（Stage-wise 新聞推薦）的中文詳解 |
| `figs/` | 從論文裁切出的圖與表（`a_` 開頭為 Paper A，`b_` 開頭為 Paper B） |
| `*.pdf` | 兩篇原始論文（CC BY 4.0） |

## 投影片大綱

1. 標題
2. 為什麼選這兩篇論文
3. Paper A：問題與動機
4. Paper A：LISRec 流程
5. Paper A：案例分析
6. Paper A：實驗結果
7. Paper B：問題與動機
8. Paper B：模型架構
9. Paper B：訓練目標與實驗設定
10. Paper B：實驗結果
11. 兩篇比較
12. 總結與討論
13. 感謝與 Q&A

## 使用方式

```bash
pip install pymupdf python-pptx pillow
python make_slides.py
```

執行後會：

1. 依照 `CROPS` 設定，從 PDF 裁切圖表到 `figs/`
2. 讀取 `講稿.md` 中每個「## 投影片 N｜…」段落，放進對應投影片的備忘稿
3. 輸出 `paper_presentation.pptx`

修改方式：

- **改講稿**：直接編輯 `講稿.md`，重新執行腳本即可更新備忘稿
- **改投影片內容**：編輯 `make_slides.py` 中的 `build()`
- **改裁切範圍**：調整 `CROPS` 中的座標（PDF points，頁面大小 612 × 792）

> 若 PowerPoint 正開著 `paper_presentation.pptx`，腳本會因檔案被鎖定而無法寫入，請先關閉再執行。

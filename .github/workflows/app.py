import os
import re
import sqlite3
from collections import Counter
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import openpyxl
import docx
from pypdf import PdfReader

DB_FILE = "light_index.db"

# 要点抽出で優先度を高くする「結論・回答シグナルワード」
SIGNAL_WORDS = ["回答", "結論", "判断", "結果", "対応", "不可", "可能", "不要", "必要", "原則", "特例", "承認", "却下", "につき", "ため、", "とおり"]
# 関連ワード提案から除外する一般的な単語（ストップワード）
STOP_WORDS = {"こと", "もの", "ため", "これ", "それ", "あり", "なし", "する", "なる", "いる", "ある", "れる", "られる", "について", "における", "および", "または", "場合", "確認", "資料", "照会", "シート", "ページ"}

class UltraLightSearchApp:
    def __init__(self, root):
        self.root = root
        self.root.title("超軽量・照会資料 検索＆提案システム (メモリ消費30MB)")
        self.root.geometry("900x700")

        # --- 1. 上部：フォルダ選択＆インデックス更新 ---
        frame_top = tk.Frame(root, pady=5)
        frame_top.pack(fill=tk.X, padx=10)
        tk.Button(frame_top, text="① 資料フォルダを選択して記憶（初回・更新時）", command=self.build_index, bg="#e1f5fe").pack(side=tk.LEFT, padx=5)
        self.lbl_status = tk.Label(frame_top, text="状態: 準備完了 (超軽量モード)", fg="green")
        self.lbl_status.pack(side=tk.LEFT, padx=10)

        # --- 2. 検索バー ---
        frame_search = tk.LabelFrame(root, text="② キーワード検索（スペース区切りでAND検索）", pady=5, padx=5)
        frame_search.pack(fill=tk.X, padx=10, pady=5)

        self.ent_keyword = tk.Entry(frame_search, font=("Meiryo", 11))
        self.ent_keyword.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)
        self.ent_keyword.bind("<Return>", lambda event: self.run_search())

        tk.Button(frame_search, text="🔍 瞬時に検索＆提案", command=self.run_search, bg="#c8e6c9", font=("Meiryo", 10, "bold")).pack(side=tk.LEFT, padx=5)

        # --- 3. 提案エリア（関連ワード＆要点自動ピックアップ） ---
        frame_suggest = tk.LabelFrame(root, text="💡 システムからの提案（関連キーワード ＆ 要点ピックアップ）", pady=5, padx=5, bg="#fffde7")
        frame_suggest.pack(fill=tk.X, padx=10, pady=5)

        self.lbl_related = tk.Label(frame_suggest, text="【関連ワード提案】: 検索すると、よく一緒に使われる単語がここに表示されます（クリックで追加検索）", bg="#fffde7", anchor="w", justify="left", wraplength=850, fg="#0d47a1", font=("Meiryo", 9, "bold"))
        self.lbl_related.pack(fill=tk.X, padx=5, pady=2)

        self.frame_chips = tk.Frame(frame_suggest, bg="#fffde7")
        self.frame_chips.pack(fill=tk.X, padx=5, pady=2)

        self.txt_summary = scrolledtext.ScrolledText(frame_suggest, height=6, bg="#fff9c4", font=("Meiryo", 9))
        self.txt_summary.pack(fill=tk.X, padx=5, pady=3)

        # --- 4. 検索結果の詳細プレビュー（ファイルを開かずに確認） ---
        tk.Label(root, text="📄 該当資料の抜粋プレビュー（ダブルクリック不要で中身を確認）", font=("Meiryo", 9, "bold")).pack(anchor="w", padx=10)
        self.txt_results = scrolledtext.ScrolledText(root, height=15, font=("Meiryo", 10))
        self.txt_results.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        self.txt_results.tag_config("highlight", background="#ffff00", foreground="#d50000")
        self.txt_results.tag_config("file_header", background="#e0e0e0", font=("Meiryo", 10, "bold"))

    def extract_text(self, filepath):
        ext = os.path.splitext(filepath)[1].lower()
        chunks = []
        try:
            if ext == ".xlsx":
                wb = openpyxl.load_workbook(filepath, data_only=True, read_only=True)
                for sheet in wb.sheetnames:
                    s = wb[sheet]
                    sheet_lines = []
                    for row in s.iter_rows(values_only=True):
                        row_txt = " | ".join([str(cell).strip() for cell in row if cell is not None and str(cell).strip() != ""])
                        if row_txt:
                            sheet_lines.append(row_txt)
                    if sheet_lines:
                        full_sheet_txt = "\n".join(sheet_lines)
                        # 800文字ごとに分割して登録
                        for i in range(0, len(full_sheet_txt), 700):
                            chunks.append((f"シート:{sheet}", full_sheet_txt[i:i+800]))
                wb.close()
            elif ext == ".docx":
                doc = docx.Document(filepath)
                full_txt = "\n".join([p.text.strip() for p in doc.paragraphs if p.text.strip()])
                for i in range(0, len(full_txt), 700):
                    chunks.append(("Word本文", full_txt[i:i+800]))
            elif ext == ".pdf":
                reader = PdfReader(filepath)
                for idx, page in enumerate(reader.pages):
                    t = page.extract_text()
                    if t and t.strip():
                        chunks.append((f"P.{idx+1}", t.strip()))
            elif ext in [".txt", ".csv"]:
                with open(filepath, "r", encoding="utf-8", errors="ignore") as f:
                    full_txt = f.read()
                    for i in range(0, len(full_txt), 700):
                        chunks.append(("テキスト", full_txt[i:i+800]))
        except Exception:
            pass
        return chunks

    def build_index(self):
        folder = filedialog.askdirectory(title="照会関係資料フォルダを選択")
        if not folder:
            return
        self.lbl_status.config(text="状態: 高速インデックス作成中...", fg="blue")
        self.root.update()

        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()
        c.execute("DROP TABLE IF EXISTS docs")
        c.execute("CREATE VIRTUAL TABLE docs USING fts5(filename, location, filepath, content, tokenize='trigram')")

        file_count = 0
        for root_dir, _, files in os.walk(folder):
            for f in files:
                if f.startswith("~$"):
                    continue
                if f.lower().endswith((".xlsx", ".docx", ".pdf", ".txt", ".csv")):
                    fullpath = os.path.join(root_dir, f)
                    chunks = self.extract_text(fullpath)
                    if chunks:
                        for loc, content in chunks:
                            c.execute("INSERT INTO docs VALUES (?, ?, ?, ?)", (f, loc, fullpath, content))
                        file_count += 1

        conn.commit()
        conn.close()
        self.lbl_status.config(text=f"状態: 記憶完了 ({file_count}ファイル)", fg="green")
        messagebox.showinfo("完了", f"{file_count} 個のファイルを高速データベースに記憶しました。")

    def add_keyword_and_search(self, word):
        current = self.ent_keyword.get().strip()
        if word not in current:
            self.ent_keyword.delete(0, tk.END)
            self.ent_keyword.insert(0, f"{current} {word}".strip())
            self.run_search()

    def analyze_and_propose(self, keywords, rows):
        """AIを使わずに0.05秒で『関連キーワード提案』と『重要文（結論）の抽出』を行うアルゴリズム"""
        all_text = "\n".join([r[3] for r in rows])

        # 1. 頻出する漢字・カタカナ・英数字の熟語（2〜8文字）を正規表現で高速抽出し、関連ワードを提案
        candidates = re.findall(r'[一-龥ァ-ヶーa-zA-Z0-9]{2,8}', all_text)
        filtered = [
            w for w in candidates
            if w not in STOP_WORDS
            and not any(w in kw or kw in w for kw in keywords)
            and not w.isdigit()
        ]
        most_common = [w for w, _ in Counter(filtered).most_common(8)]

        # 提案ボタン（チップ）の再描画
        for widget in self.frame_chips.winfo_children():
            widget.destroy()
        if most_common:
            tk.Label(self.frame_chips, text="クリックで絞り込み 👉 ", bg="#fffde7", font=("Meiryo", 9)).pack(side=tk.LEFT)
            for word in most_common:
                btn = tk.Button(self.frame_chips, text=f"＋ {word}", bg="#e3f2fd", relief=tk.FLAT,
                                command=lambda w=word: self.add_keyword_and_search(w))
                btn.pack(side=tk.LEFT, padx=3)

        # 2. 結論・回答らしき重要文をスコアリングして上位4文を自動ピックアップ
        scored_sentences = []
        seen = set()
        for fname, loc, _, content in rows:
            # 改行や句点で文を分割
            lines = re.split(r'[\n。]', content)
            for line in lines:
                line = line.strip()
                if len(line) < 12 or len(line) > 180 or line in seen:
                    continue
                seen.add(line)

                # スコア計算：検索キーワードの含有数 ＋ 結論シグナルワードの含有数
                kw_score = sum(3 for kw in keywords if kw.lower() in line.lower())
                sig_score = sum(2 for sig in SIGNAL_WORDS if sig in line)
                total_score = kw_score + sig_score

                if kw_score > 0 and total_score >= 4:
                    scored_sentences.append((total_score, f"📌 [{fname} ({loc})] {line}"))

        scored_sentences.sort(key=lambda x: x[0], reverse=True)
        self.txt_summary.delete("1.0", tk.END)
        if scored_sentences:
            self.txt_summary.insert(tk.END, "【自動抽出された重要ポイント・回答候補】\n")
            for _, text in scored_sentences[:5]:
                self.txt_summary.insert(tk.END, text + "\n")
        else:
            self.txt_summary.insert(tk.END, "※ ヒットした資料の中から、結論・回答に関する定型文は検出されませんでした。下のプレビューを参照してください。")

    def run_search(self):
        kw_input = self.ent_keyword.get().strip()
        if not kw_input:
            return
        if not os.path.exists(DB_FILE):
            messagebox.showwarning("未準備", "先に「① 資料フォルダを選択して記憶」ボタンを押してください。")
            return

        keywords = [w for w in kw_input.split() if len(w) > 0]
        conn = sqlite3.connect(DB_FILE)
        c = conn.cursor()

        # 3文字以上の単語は高速なFTS5検索、2文字以下の単語はLIKE検索を組み合わせるハイブリッド方式
        fts_words = [f'"{w}"' for w in keywords if len(w) >= 3]
        if fts_words:
            query = " AND ".join(fts_words)
            rows = c.execute("SELECT filename, location, filepath, content FROM docs WHERE docs MATCH ? LIMIT 30", (query,)).fetchall()
            # 2文字以下のキーワードがある場合は追加フィルタ
            for w in keywords:
                if len(w) < 3:
                    rows = [r for r in rows if w.lower() in r[3].lower()]
        else:
            # すべて2文字以下の場合（例：「旅費」など）
            like_clauses = " AND ".join(["content LIKE ?" for _ in keywords])
            params = [f"%{w}%" for w in keywords]
            rows = c.execute(f"SELECT filename, location, filepath, content FROM docs WHERE {like_clauses} LIMIT 30", params).fetchall()

        conn.close()

        self.txt_results.delete("1.0", tk.END)
        if not rows:
            self.txt_summary.delete("1.0", tk.END)
            self.txt_results.insert(tk.END, "該当する資料が見つかりませんでした。")
            return

        # 関連ワード提案と要点ピックアップの実行
        self.analyze_and_propose(keywords, rows)

        # 検索結果の表示
        for idx, (fname, loc, fpath, content) in enumerate(rows, 1):
            header = f"■ {idx}. {fname}  【場所: {loc}】  (パス: {fpath})\n"
            self.txt_results.insert(tk.END, header, "file_header")
            self.txt_results.insert(tk.END, f"{content.strip()}\n" + ("-"*80) + "\n\n")

        # 検索キーワードを黄色＋赤文字でハイライト表示
        for kw in keywords:
            start_pos = "1.0"
            while True:
                start_pos = self.txt_results.search(kw, start_pos, stopindex=tk.END, nocase=True)
                if not start_pos:
                    break
                end_pos = f"{start_pos}+{len(kw)}c"
                self.txt_results.tag_add("highlight", start_pos, end_pos)
                start_pos = end_pos

if __name__ == "__main__":
    root = tk.Tk()
    app = UltraLightSearchApp(root)
    root.mainloop()

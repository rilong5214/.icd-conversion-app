import os
import subprocess
import tkinter as tk
from tkinter import messagebox, filedialog, ttk
from tkinterdnd2 import DND_FILES, TkinterDnD

class ICDConvertApp:
    def __init__(self, root):
        self.root = root
        self.root.title("ICD変換ツール")
        self.root.geometry("600x410")
        self.icd_files = []
        self.mode = tk.StringVar(value="stp")
        self.output_dir = tk.StringVar(value="")  # 出力先フォルダ（未選択時は空欄）

        # 変換モード選択
        mode_frame = tk.Frame(root)
        mode_frame.pack(pady=5)
        tk.Label(mode_frame, text="変換モード：", font=("Arial", 11)).pack(side=tk.LEFT)
        tk.Radiobutton(mode_frame, text="ICD→STEP", variable=self.mode, value="stp", font=("Arial", 11)).pack(side=tk.LEFT)
        tk.Radiobutton(mode_frame, text="ICD→Parasolid", variable=self.mode, value="parasolid", font=("Arial", 11)).pack(side=tk.LEFT)

        # 出力先フォルダ選択
        output_frame = tk.Frame(root)
        output_frame.pack(pady=4)
        tk.Label(output_frame, text="出力先フォルダ：", font=("Arial", 10)).pack(side=tk.LEFT)
        self.output_dir_label = tk.Label(output_frame, text="（未選択時は入力元と同じ場所）", font=("Arial", 9), fg="gray")
        self.output_dir_label.pack(side=tk.LEFT, padx=4)
        tk.Button(output_frame, text="出力先を選択", font=("Arial", 9), command=self.select_output_dir).pack(side=tk.LEFT, padx=8)
        tk.Button(output_frame, text="選択解除", font=("Arial", 9), command=self.clear_output_dir).pack(side=tk.LEFT)

        # ドロップ案内
        self.label = tk.Label(root, text="ここに .icd ファイルをドロップしてください", font=("Arial", 12))
        self.label.pack(pady=6)

        # ファイルリスト
        self.listbox = tk.Listbox(root, selectmode=tk.BROWSE, width=80, height=8)
        self.listbox.pack(padx=10, pady=8)
        self.listbox.drop_target_register(DND_FILES)
        self.listbox.dnd_bind('<<Drop>>', self.on_drop)

        # プログレスバー
        self.progress = ttk.Progressbar(root, orient="horizontal", length=560, mode="determinate")
        self.progress.pack(pady=6)
        self.progress["value"] = 0

        # ボタン
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=8)
        self.convert_btn = tk.Button(btn_frame, text="変換", width=12, command=self.convert_files)
        self.convert_btn.pack(side=tk.LEFT, padx=10)
        self.clear_btn = tk.Button(btn_frame, text="クリア", width=12, command=self.clear_files)
        self.clear_btn.pack(side=tk.LEFT, padx=10)

    def select_output_dir(self):
        folder = filedialog.askdirectory(title="出力先フォルダを選択")
        if folder:
            self.output_dir.set(folder)
            self.output_dir_label.config(text=folder, fg="black")

    def clear_output_dir(self):
        self.output_dir.set("")
        self.output_dir_label.config(text="（未選択時は入力元と同じ場所）", fg="gray")

    def on_drop(self, event):
        files = self.root.tk.splitlist(event.data)
        added = 0
        for f in files:
            if f.lower().endswith(".icd") and f not in self.icd_files:
                self.icd_files.append(f)
                self.listbox.insert(tk.END, os.path.basename(f))
                added += 1
        if added == 0:
            messagebox.showinfo("情報", "新しい .icd ファイルはありません。")

    def convert_files(self):
        if not self.icd_files:
            messagebox.showwarning("警告", "変換対象ファイルがありません。")
            return
        mode = self.mode.get()
        exe_path = r'C:\ICADSX\BIN\ICD2STP.exe' if mode == "stp" else r'C:\ICADSX\BIN\ICD2PS.exe'

        total = len(self.icd_files)
        self.progress["maximum"] = total
        self.progress["value"] = 0

        output_folder = self.output_dir.get().strip()
        success = 0

        for idx, icd_path in enumerate(self.icd_files, 1):
            base_name = os.path.splitext(os.path.basename(icd_path))[0]
            # 出力先フォルダが指定されていればそこに、なければ入力元と同じ
            out_dir = output_folder if output_folder else os.path.dirname(icd_path)
            out_path = os.path.join(out_dir, base_name)
            icd_path_abs = os.path.abspath(icd_path)
            out_path_abs = os.path.abspath(out_path)
            command = [
                exe_path,
                '-ls',
                f'-i"{icd_path_abs}"',
                f'-o"{out_path_abs}"'
            ]
            command_str = ' '.join(command)
            try:
                subprocess.run(command_str, shell=True, check=True)
                # 最初の1件だけ自動でフォルダを開く（全件開くのは煩雑なため）
                if idx == 1:
                    os.startfile(out_dir)
                success += 1
            except subprocess.CalledProcessError as e:
                messagebox.showerror("変換失敗", f"{os.path.basename(icd_path)} の変換に失敗しました。\n\n{e}")
            self.progress["value"] = idx
            self.root.update_idletasks()  # プログレスバー更新

        messagebox.showinfo("完了", f"{success} ファイルの変換が完了しました。")
        self.progress["value"] = 0  # リセット

    def clear_files(self):
        self.icd_files = []
        self.listbox.delete(0, tk.END)
        self.progress["value"] = 0

if __name__ == "__main__":
    root = TkinterDnD.Tk()
    app = ICDConvertApp(root)
    root.mainloop()

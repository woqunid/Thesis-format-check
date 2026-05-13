from pathlib import Path
from tkinter import BOTH, END, LEFT, RIGHT, Button, Frame, Label, Tk, filedialog, messagebox
from tkinter.scrolledtext import ScrolledText

from format_checker.report import render_report
from format_checker.rules import check_file

WINDOW_SIZE = "980x680"
TEXT_HEIGHT = 34


class FormatCheckerApp:
    def __init__(self, root: Tk) -> None:
        self.root = root
        self.selected_file: Path | None = None
        self.report_text = ""
        self.file_label = Label(root, text="未选择文件", anchor="w")
        self.output = ScrolledText(root, height=TEXT_HEIGHT, wrap="word")
        self._build()

    def _build(self) -> None:
        self.root.title("毕业论文格式检测器")
        self.root.geometry(WINDOW_SIZE)

        top_bar = Frame(self.root)
        top_bar.pack(fill="x", padx=12, pady=10)

        Button(top_bar, text="选择 Word 文件", command=self.choose_file).pack(side=LEFT)
        Button(top_bar, text="开始检测", command=self.run_check).pack(side=LEFT, padx=8)
        Button(top_bar, text="保存报告", command=self.save_report).pack(side=RIGHT)

        self.file_label.pack(fill="x", padx=12)
        self.output.pack(fill=BOTH, expand=True, padx=12, pady=10)
        self._set_output(_intro_text())

    def choose_file(self) -> None:
        filename = filedialog.askopenfilename(
            title="选择 Word 文档",
            filetypes=[("Word 文档", "*.docx *.doc"), ("所有文件", "*.*")],
        )
        if not filename:
            return
        self.selected_file = Path(filename)
        self.file_label.config(text=str(self.selected_file))

    def run_check(self) -> None:
        if self.selected_file is None:
            messagebox.showwarning("未选择文件", "请先选择 .docx 或 .doc 文件。")
            return
        try:
            violations = check_file(self.selected_file)
            self.report_text = render_report(violations)
            self._set_output(self.report_text)
        except Exception as exc:
            self.report_text = f"检测失败：{exc}"
            self._set_output(self.report_text)

    def save_report(self) -> None:
        if not self.report_text:
            messagebox.showwarning("没有报告", "请先完成检测。")
            return
        filename = filedialog.asksaveasfilename(
            title="保存检测报告",
            defaultextension=".txt",
            filetypes=[("文本文件", "*.txt")],
        )
        if not filename:
            return
        Path(filename).write_text(self.report_text, encoding="utf-8")

    def _set_output(self, text: str) -> None:
        self.output.delete("1.0", END)
        self.output.insert("1.0", text)


def main() -> None:
    root = Tk()
    FormatCheckerApp(root)
    root.mainloop()


def _intro_text() -> str:
    return "\n".join(
        [
            "请选择 .docx 文件后点击“开始检测”。",
            "",
            "检测范围：字体、字号、加粗、对齐、缩进、固定行距、关键词分隔符、页眉等可解析项目。",
            "不检测：目录页码准确性、图片实际位置、表格实际相邻关系、分页位置等 python-docx 无法可靠解析的项目。",
            "",
            ".doc 文件会明确提示先转换为 .docx。",
        ]
    )

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
from report_generator import generate_lateness_report


class LatenessReportApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Lateness Report Generator")
        self.root.geometry("620x420")
        self.root.resizable(False, False)

        style = ttk.Style()
        style.configure("TButton", padding=6)
        style.configure("TLabel", padding=4)
        style.configure("Header.TLabel", font=("Arial", 14, "bold"))
        style.configure("Big.TButton", font=("Arial", 12, "bold"), padding=10)

        main_frame = ttk.Frame(root, padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Attendance \u2192 Lateness Report", style="Header.TLabel").pack(pady=(0, 15))

        # File selection
        file_frame = ttk.LabelFrame(main_frame, text="Input File (.xls / .xlsx)", padding=10)
        file_frame.pack(fill=tk.X, pady=5)

        self.file_path_var = tk.StringVar()
        ttk.Entry(file_frame, textvariable=self.file_path_var, width=55).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(file_frame, text="Browse...", command=self.browse_file).pack(side=tk.LEFT)

        # Top N
        opt_frame = ttk.LabelFrame(main_frame, text="Options", padding=10)
        opt_frame.pack(fill=tk.X, pady=5)

        ttk.Label(opt_frame, text="Top N employees:").pack(side=tk.LEFT)
        self.top_n_var = tk.IntVar(value=10)
        ttk.Spinbox(opt_frame, from_=1, to=100, textvariable=self.top_n_var, width=5).pack(side=tk.LEFT, padx=5)

        # Output
        out_frame = ttk.LabelFrame(main_frame, text="Output File", padding=10)
        out_frame.pack(fill=tk.X, pady=5)

        self.output_path_var = tk.StringVar()
        ttk.Entry(out_frame, textvariable=self.output_path_var, width=55).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(out_frame, text="Browse...", command=self.browse_output).pack(side=tk.LEFT)

        # ── Generate button (big & centered) ──
        gen_frame = ttk.Frame(main_frame)
        gen_frame.pack(fill=tk.X, pady=(20, 5))

        self.generate_btn = tk.Button(
            gen_frame, text="Generate Report", command=self.generate,
            font=("Arial", 13, "bold"), bg="#4472C4", fg="white",
            activebackground="#2F5496", activeforeground="white",
            relief="raised", bd=2, padx=30, pady=8, cursor="hand2",
        )
        self.generate_btn.pack()

        # Progress + status
        status_frame = ttk.Frame(main_frame)
        status_frame.pack(fill=tk.X, pady=(5, 0))

        self.progress = ttk.Progressbar(status_frame, mode="indeterminate", length=300)
        self.progress.pack(side=tk.LEFT, padx=(0, 10))

        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(status_frame, textvariable=self.status_var).pack(side=tk.LEFT)

    def browse_file(self):
        path = filedialog.askopenfilename(
            title="Select Attendance File",
            filetypes=[("Excel files", "*.xls *.xlsx *.xlsm"), ("All files", "*.*")]
        )
        if path:
            self.file_path_var.set(path)
            if not self.output_path_var.get():
                directory = os.path.dirname(path)
                self.output_path_var.set(os.path.join(directory, "Lateness_Report.xlsx"))

    def browse_output(self):
        path = filedialog.asksaveasfilename(
            title="Save Report As",
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )
        if path:
            self.output_path_var.set(path)

    def generate(self):
        input_path = self.file_path_var.get().strip()
        output_path = self.output_path_var.get().strip()
        top_n = self.top_n_var.get()

        if not input_path:
            messagebox.showwarning("Warning", "Please select an input file.")
            return
        if not os.path.isfile(input_path):
            messagebox.showerror("Error", "Input file not found.")
            return
        if not output_path:
            messagebox.showwarning("Warning", "Please specify an output file.")
            return

        self.generate_btn.config(state="disabled", text="Generating...")
        self.progress.start()
        self.status_var.set("Generating...")

        def run():
            try:
                generate_lateness_report(input_path, output_path, top_n, self.update_status)
                self.root.after(0, lambda: self._done(True))
            except Exception as e:
                err = str(e)
                self.root.after(0, lambda: self._done(False, err))

        threading.Thread(target=run, daemon=True).start()

    def update_status(self, msg):
        self.root.after(0, lambda: self.status_var.set(msg))

    def _done(self, success, error=None):
        self.progress.stop()
        self.generate_btn.config(state="normal", text="Generate Report")
        if success:
            self.status_var.set("Done!")
            messagebox.showinfo("Success", f"Report saved to:\n{self.output_path_var.get()}")
        else:
            self.status_var.set("Error")
            messagebox.showerror("Error", f"Failed to generate report:\n{error}")


if __name__ == "__main__":
    root = tk.Tk()
    app = LatenessReportApp(root)
    root.mainloop()

import tkinter as tk
from tkinter import ttk, messagebox
import threading, sys, io, subprocess
from datetime import date
from pathlib import Path

MONTHS = ["January","February","March","April","May","June",
          "July","August","September","October","November","December"]
today = date.today()


class ScraperApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Attendance Scraper — Indian Inovatix")
        self.resizable(False, False)
        self.configure(bg="#f0f2f5")
        self._build_ui()
        self._center(520, 500)

    def _center(self, w, h):
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _build_ui(self):
        hdr = tk.Frame(self, bg="#1a1a2e", pady=14)
        hdr.pack(fill="x")
        tk.Label(hdr, text="Attendance Data Scraper", font=("Segoe UI", 15, "bold"),
                 fg="white", bg="#1a1a2e").pack()
        tk.Label(hdr, text="Indian Inovatix Limited — HR Portal",
                 font=("Segoe UI", 9), fg="#aaaacc", bg="#1a1a2e").pack()

        ctrl = tk.Frame(self, bg="#f0f2f5", pady=18)
        ctrl.pack(fill="x", padx=30)

        tk.Label(ctrl, text="Month", font=("Segoe UI", 10, "bold"),
                 bg="#f0f2f5").grid(row=0, column=0, sticky="w", padx=(0, 10))
        self.month_var = tk.StringVar(value=MONTHS[today.month - 1])
        ttk.Combobox(ctrl, textvariable=self.month_var, values=MONTHS,
                     state="readonly", font=("Segoe UI", 10), width=14
                     ).grid(row=0, column=1, padx=(0, 20))

        tk.Label(ctrl, text="Year", font=("Segoe UI", 10, "bold"),
                 bg="#f0f2f5").grid(row=0, column=2, sticky="w", padx=(0, 10))
        self.year_var = tk.StringVar(value=str(today.year))
        ttk.Combobox(ctrl, textvariable=self.year_var,
                     values=[str(y) for y in range(today.year - 2, today.year + 1)],
                     state="readonly", font=("Segoe UI", 10), width=8
                     ).grid(row=0, column=3)

        self.headless_var = tk.BooleanVar(value=True)
        tk.Checkbutton(ctrl, text="Run in background (headless)",
                       variable=self.headless_var,
                       font=("Segoe UI", 9), bg="#f0f2f5"
                       ).grid(row=1, column=0, columnspan=4, sticky="w", pady=(10, 0))

        self.autopush_var = tk.BooleanVar(value=False)
        tk.Checkbutton(ctrl, text="Push to GitHub after scraping",
                       variable=self.autopush_var,
                       font=("Segoe UI", 9), bg="#f0f2f5"
                       ).grid(row=2, column=0, columnspan=4, sticky="w")

        btns = tk.Frame(self, bg="#f0f2f5")
        btns.pack(pady=(0, 12))

        self.scrape_btn = tk.Button(btns, text="  Start Scraping  ",
            font=("Segoe UI", 11, "bold"), bg="#0066cc", fg="white",
            activebackground="#0052a3", activeforeground="white",
            relief="flat", padx=20, pady=8, cursor="hand2",
            command=self._start_scrape)
        self.scrape_btn.pack(side="left", padx=6)

        tk.Button(btns, text="Clear Log", font=("Segoe UI", 9),
            bg="#e0e0e0", fg="#333", activebackground="#c8c8c8",
            relief="flat", padx=12, pady=8, cursor="hand2",
            command=self._clear_log).pack(side="left", padx=6)

        self.status_var = tk.StringVar(value="Ready")
        tk.Label(self, textvariable=self.status_var,
                 font=("Segoe UI", 9, "italic"),
                 fg="#555", bg="#f0f2f5").pack()

        log_frame = tk.Frame(self, bg="#f0f2f5", padx=12, pady=6)
        log_frame.pack(fill="both", expand=True, padx=12, pady=(0, 12))
        tk.Label(log_frame, text="Log", font=("Segoe UI", 9, "bold"),
                 bg="#f0f2f5", anchor="w").pack(fill="x")

        tf = tk.Frame(log_frame)
        tf.pack(fill="both", expand=True)
        self.log = tk.Text(tf, font=("Consolas", 9), bg="#1e1e2e", fg="#cdd6f4",
                           insertbackground="white", wrap="word",
                           state="disabled", relief="flat", pady=6, padx=8)
        sb = tk.Scrollbar(tf, command=self.log.yview)
        self.log.configure(yscrollcommand=sb.set)
        sb.pack(side="right", fill="y")
        self.log.pack(side="left", fill="both", expand=True)
        self.log.tag_config("info",    foreground="#89b4fa")
        self.log.tag_config("success", foreground="#a6e3a1")
        self.log.tag_config("warn",    foreground="#f9e2af")
        self.log.tag_config("error",   foreground="#f38ba8")
        self.log.tag_config("dim",     foreground="#6c7086")

    def _log(self, text, tag="dim"):
        self.log.configure(state="normal")
        self.log.insert("end", text + "\n", tag)
        self.log.see("end")
        self.log.configure(state="disabled")

    def _clear_log(self):
        self.log.configure(state="normal")
        self.log.delete("1.0", "end")
        self.log.configure(state="disabled")
        self.status_var.set("Ready")

    def _start_scrape(self):
        month_num = MONTHS.index(self.month_var.get()) + 1
        year_num  = int(self.year_var.get())
        headless  = self.headless_var.get()
        autopush  = self.autopush_var.get()
        self.scrape_btn.configure(state="disabled", text="  Scraping...  ")
        self.status_var.set(f"Scraping {self.month_var.get()} {year_num}...")
        self._log(f"{'='*50}", "dim")
        self._log(f"  {self.month_var.get()} {year_num}  (headless={headless})", "info")
        self._log(f"{'='*50}", "dim")
        threading.Thread(target=self._run_scrape,
                         args=(year_num, month_num, headless, autopush),
                         daemon=True).start()

    def _run_scrape(self, year, month, headless, autopush=False):
        try:
            old_stdout = sys.stdout
            sys.stdout = LogRedirector(self)
            import config as _cfg
            _cfg.HEADLESS = headless
            from spine_scraper import fetch_attendance
            result = fetch_attendance(year=year, month=month)
            sys.stdout = old_stdout
            n = len(result.get("records", []))
            if "fatal_error" in result and n == 0:
                self._log(f"\nFAILED: {result['fatal_error']}", "error")
                self.after(0, self._scrape_done, False, result.get("fatal_error", ""))
            else:
                self._log(f"\nDone! {n} records saved.", "success")
                if autopush:
                    self._git_push(_cfg.DATA_FILE, year, month)
                self.after(0, self._scrape_done, True, n)
        except Exception as e:
            sys.stdout = sys.__stdout__
            self._log(f"\nERROR: {e}", "error")
            self.after(0, self._scrape_done, False, str(e))

    def _git_push(self, data_file, year, month):
        repo_dir   = str(Path(__file__).parent)
        commit_msg = f"attendance: {MONTHS[month-1]} {year}"
        self._log("\nPushing to GitHub...", "info")
        try:
            for cmd in [["git","add", data_file],
                        ["git","commit","-m", commit_msg],
                        ["git","push"]]:
                res = subprocess.run(cmd, cwd=repo_dir, capture_output=True, text=True)
                out = (res.stdout + res.stderr).strip()
                if out:
                    self._log(out, "info" if res.returncode == 0 else "warn")
                if res.returncode != 0 and cmd[1] != "commit":
                    self._log(f"Git error: {' '.join(cmd)}", "error")
                    return
            self._log("Pushed to GitHub.", "success")
        except Exception as e:
            self._log(f"Git push error: {e}", "error")

    def _scrape_done(self, success, info):
        self.scrape_btn.configure(state="normal", text="  Start Scraping  ")
        if success:
            import config as _cfg
            self.status_var.set(f"Done — {info} records saved.")
            messagebox.showinfo("Done", f"Saved {info} records.\n\nFile: {_cfg.DATA_FILE}")
        else:
            self.status_var.set("Failed — see log.")
            messagebox.showerror("Failed", str(info))


class LogRedirector(io.TextIOBase):
    def __init__(self, app):
        self.app = app

    def write(self, text):
        if text.strip():
            tag = "dim"
            if "[INFO]"  in text: tag = "info"
            if "[DONE]"  in text: tag = "success"
            if "[WARN]"  in text: tag = "warn"
            if "[ERROR]" in text or "[FATAL]" in text: tag = "error"
            if text.strip().startswith("["): tag = "info"
            self.app.after(0, self.app._log, text.rstrip(), tag)
        return len(text)

    def flush(self):
        pass


if __name__ == "__main__":
    import os
    os.chdir(Path(__file__).parent)
    ScraperApp().mainloop()

import threading
from pathlib import Path
from tkinter import filedialog
from tkinter import messagebox

import customtkinter as ctk

from app.config import ROOT_DIR, get_latest_json_url, load_version
from app.logger import logger
from app.services.business_logic import format_result_summary, process_folder_to_excel
from app.update_manager import download_update, get_update_status
from app.updater_launcher import launch_updater


class AppGUI:
    def __init__(self):
        metadata = load_version()

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.root = ctk.CTk()
        self.root.title(metadata.get("app_name", "BV Application"))
        self.root.geometry("1000x700")
        self.root.minsize(800, 520)

        self.status_var = ctk.StringVar(value="Ready")
        self.folder_var = ctk.StringVar(value="")
        self.output_var = ctk.StringVar(value=str(Path.cwd() / "quan_ly_van_ban.xlsx"))
        self.metadata = metadata
        self.update_check_running = False
        self.task_running = False

        self._build_layout(metadata)
        self._append_log("Sẵn sàng xử lý danh sách file PDF.")
        self.root.after(800, self._check_update_on_startup)

    def _build_layout(self, metadata):
        self.root.grid_columnconfigure(0, weight=1)
        self.root.grid_rowconfigure(1, weight=1)

        header = ctk.CTkFrame(self.root, corner_radius=0)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            header,
            text=metadata.get("app_name", "BV Application"),
            font=ctk.CTkFont(size=22, weight="bold"),
        )
        title.grid(row=0, column=0, sticky="w", padx=20, pady=(16, 4))

        version = ctk.CTkLabel(
            header,
            text=f"Version {metadata.get('version', '0.0.0')}",
        )
        version.grid(row=1, column=0, sticky="w", padx=20, pady=(0, 16))

        body = ctk.CTkFrame(self.root, corner_radius=0)
        body.grid(row=1, column=0, sticky="nsew", padx=20, pady=20)
        body.grid_columnconfigure(0, weight=1)
        body.grid_rowconfigure(2, weight=1)

        form = ctk.CTkFrame(body)
        form.grid(row=0, column=0, sticky="ew", pady=(0, 12))
        form.grid_columnconfigure(1, weight=1)

        folder_label = ctk.CTkLabel(form, text="Thư mục PDF")
        folder_label.grid(row=0, column=0, sticky="w", padx=12, pady=(12, 6))

        folder_entry = ctk.CTkEntry(form, textvariable=self.folder_var)
        folder_entry.grid(row=0, column=1, sticky="ew", padx=(0, 8), pady=(12, 6))

        folder_button = ctk.CTkButton(form, text="Chọn", width=90, command=self._choose_folder)
        folder_button.grid(row=0, column=2, sticky="e", padx=(0, 12), pady=(12, 6))

        output_label = ctk.CTkLabel(form, text="File Excel")
        output_label.grid(row=1, column=0, sticky="w", padx=12, pady=(6, 12))

        output_entry = ctk.CTkEntry(form, textvariable=self.output_var)
        output_entry.grid(row=1, column=1, sticky="ew", padx=(0, 8), pady=(6, 12))

        output_button = ctk.CTkButton(form, text="Lưu tại", width=90, command=self._choose_output)
        output_button.grid(row=1, column=2, sticky="e", padx=(0, 12), pady=(6, 12))

        actions = ctk.CTkFrame(body)
        actions.grid(row=1, column=0, sticky="ew", pady=(0, 12))

        self.run_button = ctk.CTkButton(actions, text="Tạo Excel", command=self._run_task)
        self.run_button.pack(side="left", padx=12, pady=12)

        update_button = ctk.CTkButton(actions, text="Check Update", command=self._check_update)
        update_button.pack(side="left", padx=(0, 12), pady=12)

        status = ctk.CTkLabel(actions, textvariable=self.status_var)
        status.pack(side="left", padx=12)

        self.output_box = ctk.CTkTextbox(body)
        self.output_box.grid(row=2, column=0, sticky="nsew")
        self.output_box.insert("end", "Nhật ký xử lý\n")
        self.output_box.configure(state="disabled")

    def _choose_folder(self):
        folder = filedialog.askdirectory(parent=self.root)
        if folder:
            self.folder_var.set(folder)

    def _choose_output(self):
        output = filedialog.asksaveasfilename(
            parent=self.root,
            defaultextension=".xlsx",
            filetypes=[("Excel workbook", "*.xlsx")],
            initialfile="quan_ly_van_ban.xlsx",
        )
        if output:
            self.output_var.set(output)

    def _run_task(self):
        if self.task_running:
            return

        folder_path = self.folder_var.get().strip()
        output_path = self.output_var.get().strip()

        self.task_running = True
        self.run_button.configure(state="disabled")
        self.status_var.set("Đang xử lý...")
        self._append_log("Bắt đầu xử lý.")

        thread = threading.Thread(target=self._run_processing_task, args=(folder_path, output_path), daemon=True)
        thread.start()

    def _run_processing_task(self, folder_path, output_path):
        try:
            result = process_folder_to_excel(folder_path, output_path, self._append_log_threadsafe)
            self.root.after(0, self._finish_processing_task, result)
        except Exception as exc:
            logger.exception("Processing failed")
            self.root.after(0, self._show_processing_error, str(exc))

    def _finish_processing_task(self, result):
        self.task_running = False
        self.run_button.configure(state="normal")
        self.status_var.set("Hoàn tất" if result.ok else "Có lỗi")
        summary = format_result_summary(result)
        logger.info(summary)
        self._append_log(summary)

        if not result.ok:
            messagebox.showerror("Xử lý thất bại", result.message, parent=self.root)

    def _show_processing_error(self, error):
        self.task_running = False
        self.run_button.configure(state="normal")
        message = f"Xử lý thất bại: {error}"
        self.status_var.set("Có lỗi")
        self._append_log(message)
        messagebox.showerror("Xử lý thất bại", message, parent=self.root)

    def _check_update_on_startup(self):
        self._start_update_check(is_manual=False)

    def _check_update(self):
        self._start_update_check(is_manual=True)

    def _start_update_check(self, is_manual):
        if self.update_check_running:
            return

        self.update_check_running = True
        if is_manual:
            self.status_var.set("Checking for updates...")
            self._append_log("Checking for updates...")

        thread = threading.Thread(target=self._run_update_check, args=(is_manual,), daemon=True)
        thread.start()

    def _run_update_check(self, is_manual):
        current_version = self.metadata.get("version", "0.0.0")
        latest_json_url = get_latest_json_url(self.metadata)

        if not latest_json_url:
            message = "Update check skipped: latest.json URL is not configured"
            logger.warning(message)
        else:
            result = get_update_status(current_version, latest_json_url)
            message = result.get("message", "Update check finished")

        self.root.after(0, self._finish_update_check, result if latest_json_url else {"status": "missing_config", "message": message}, is_manual)

    def _finish_update_check(self, result, is_manual):
        self.update_check_running = False
        message = result.get("message", "Update check finished")
        self.status_var.set(message)
        logger.info("Update check: %s", message)

        if is_manual or result.get("status") == "available":
            self._append_log(message)

        if result.get("status") == "available":
            self._prompt_update(result.get("latest", {}))

    def _prompt_update(self, latest_data):
        latest_version = latest_data.get("version", "unknown")
        changelog = self._format_changelog(latest_data.get("changelog", []))
        prompt = (
            f"A new version is available: {latest_version}\n\n"
            f"What's new:\n{changelog}\n\n"
            "Do you want to update now?"
        )

        if messagebox.askyesno("Update Available", prompt, parent=self.root):
            self._download_and_launch_update(latest_data)

    def _download_and_launch_update(self, latest_data):
        download_url = latest_data.get("download_url")
        if not download_url:
            message = "Update skipped: download URL is missing"
            logger.warning(message)
            self.status_var.set(message)
            self._append_log(message)
            return

        self.status_var.set("Downloading update...")
        self._append_log("Downloading update...")
        thread = threading.Thread(target=self._run_update_download, args=(download_url,), daemon=True)
        thread.start()

    def _run_update_download(self, download_url):
        try:
            new_exe_path = download_update(
                download_url,
                ROOT_DIR / "temp",
                self.metadata.get("exe_name", "BV-App.exe"),
            )
            self.root.after(0, self._launch_downloaded_update, new_exe_path)
        except Exception as exc:
            self.root.after(0, self._show_update_download_error, str(exc))

    def _launch_downloaded_update(self, new_exe_path):
        try:
            launch_updater(new_exe_path)
            message = "Update downloaded. Closing app to apply update..."
            self.status_var.set(message)
            self._append_log(message)
            self.root.after(500, self.root.destroy)
        except Exception as exc:
            self._show_update_download_error(str(exc))

    def _show_update_download_error(self, error):
        message = f"Update failed: {error}"
        logger.warning(message)
        self.status_var.set(message)
        self._append_log(message)
        messagebox.showerror("Update Failed", message, parent=self.root)

    def _format_changelog(self, changelog):
        if isinstance(changelog, list):
            items = [str(item).strip() for item in changelog if str(item).strip()]
        elif changelog:
            items = [str(changelog).strip()]
        else:
            items = ["No changelog provided."]

        return "\n".join(f"- {item}" for item in items[:3])

    def _append_log(self, message):
        self.output_box.configure(state="normal")
        self.output_box.insert("end", f"{message}\n")
        self.output_box.see("end")
        self.output_box.configure(state="disabled")

    def _append_log_threadsafe(self, message):
        self.root.after(0, self._append_log, message)

    def run(self):
        self.root.mainloop()

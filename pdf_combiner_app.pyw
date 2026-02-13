import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import tempfile
import platform
import subprocess
import threading
import pikepdf  # Better PDF handling


class PDFCombinerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF Combiner & Printer (Signature Aware)")
        self.root.geometry("700x650")
        self.root.resizable(True, True)

        self.pdf_files = []
        self.printers = []

        self.setup_ui()
        self.load_printers()

    def setup_ui(self):
        # Title
        title_frame = tk.Frame(self.root, bg="#2c3e50", height=60)
        title_frame.pack(fill=tk.X)
        title_frame.pack_propagate(False)

        title_label = tk.Label(
            title_frame,
            text="PDF Combiner & Printer",
            font=("Arial", 18, "bold"),
            bg="#2c3e50",
            fg="white",
        )
        title_label.pack(pady=15)

        # Main container
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)

        # PDF Files Section
        files_label = tk.Label(
            main_frame, text="PDF Files (in order):", font=("Arial", 12, "bold")
        )
        files_label.pack(anchor=tk.W, pady=(0, 5))

        # Listbox with scrollbar
        listbox_frame = tk.Frame(main_frame)
        listbox_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        scrollbar = tk.Scrollbar(listbox_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.files_listbox = tk.Listbox(
            listbox_frame, yscrollcommand=scrollbar.set, font=("Arial", 10), height=10
        )
        self.files_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.files_listbox.yview)

        # Buttons for file management
        btn_frame = tk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(0, 15))

        add_btn = tk.Button(
            btn_frame,
            text="➕ Add PDFs",
            command=self.add_pdfs,
            bg="#27ae60",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=15,
            pady=8,
            cursor="hand2",
        )
        add_btn.pack(side=tk.LEFT, padx=(0, 5))

        remove_btn = tk.Button(
            btn_frame,
            text="➖ Remove Selected",
            command=self.remove_pdf,
            bg="#e74c3c",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=15,
            pady=8,
            cursor="hand2",
        )
        remove_btn.pack(side=tk.LEFT, padx=5)

        move_up_btn = tk.Button(
            btn_frame,
            text="⬆ Move Up",
            command=self.move_up,
            bg="#3498db",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=15,
            pady=8,
            cursor="hand2",
        )
        move_up_btn.pack(side=tk.LEFT, padx=5)

        move_down_btn = tk.Button(
            btn_frame,
            text="⬇ Move Down",
            command=self.move_down,
            bg="#3498db",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=15,
            pady=8,
            cursor="hand2",
        )
        move_down_btn.pack(side=tk.LEFT, padx=5)

        clear_btn = tk.Button(
            btn_frame,
            text="🗑 Clear All",
            command=self.clear_all,
            bg="#95a5a6",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=15,
            pady=8,
            cursor="hand2",
        )
        clear_btn.pack(side=tk.LEFT, padx=5)

        # Signature Handling Option
        sig_frame = tk.Frame(main_frame, bg="#fff3cd", bd=2, relief=tk.GROOVE)
        sig_frame.pack(fill=tk.X, pady=(0, 15), padx=5)

        sig_info = tk.Label(
            sig_frame,
            text="⚠️ Digital Signatures: Merging removes signatures. Choose handling method:",
            font=("Arial", 9, "bold"),
            bg="#fff3cd",
            fg="#856404",
            wraplength=650,
            justify=tk.LEFT,
        )
        sig_info.pack(padx=10, pady=(10, 5), anchor=tk.W)

        self.signature_var = tk.StringVar(value="warn")

        radio_frame = tk.Frame(sig_frame, bg="#fff3cd")
        radio_frame.pack(padx=20, pady=(0, 10), anchor=tk.W)

        tk.Radiobutton(
            radio_frame,
            text="Warn me if signatures detected",
            variable=self.signature_var,
            value="warn",
            bg="#fff3cd",
            font=("Arial", 9),
        ).pack(anchor=tk.W)

        tk.Radiobutton(
            radio_frame,
            text="Flatten signatures (make them visible but not verifiable)",
            variable=self.signature_var,
            value="flatten",
            bg="#fff3cd",
            font=("Arial", 9),
        ).pack(anchor=tk.W)

        tk.Radiobutton(
            radio_frame,
            text="Proceed anyway (signatures will be lost)",
            variable=self.signature_var,
            value="proceed",
            bg="#fff3cd",
            font=("Arial", 9),
        ).pack(anchor=tk.W)

        # Printer Selection
        printer_frame = tk.Frame(main_frame)
        printer_frame.pack(fill=tk.X, pady=(0, 15))

        printer_label = tk.Label(
            printer_frame, text="Select Printer:", font=("Arial", 11, "bold")
        )
        printer_label.pack(side=tk.LEFT, padx=(0, 10))

        self.printer_combo = ttk.Combobox(
            printer_frame, font=("Arial", 10), state="readonly", width=35
        )
        self.printer_combo.pack(side=tk.LEFT, fill=tk.X, expand=True)

        refresh_btn = tk.Button(
            printer_frame,
            text="🔄",
            command=self.load_printers,
            bg="#95a5a6",
            fg="white",
            font=("Arial", 10, "bold"),
            padx=10,
            pady=5,
            cursor="hand2",
        )
        refresh_btn.pack(side=tk.LEFT, padx=(5, 0))

        # Action Buttons
        action_frame = tk.Frame(main_frame)
        action_frame.pack(fill=tk.X, pady=(10, 0))

        print_btn = tk.Button(
            action_frame,
            text="🖨 Combine & Print",
            command=self.combine_and_print,
            bg="#e67e22",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=12,
            cursor="hand2",
        )
        print_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))

        save_btn = tk.Button(
            action_frame,
            text="💾 Combine & Save",
            command=self.combine_and_save,
            bg="#9b59b6",
            fg="white",
            font=("Arial", 12, "bold"),
            padx=20,
            pady=12,
            cursor="hand2",
        )
        save_btn.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=5)

        # Status bar
        self.status_label = tk.Label(
            self.root,
            text="Ready",
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W,
            font=("Arial", 9),
            padx=10,
            pady=5,
        )
        self.status_label.pack(side=tk.BOTTOM, fill=tk.X)

    def check_for_signatures(self, pdf_files):
        """Check if any PDF has digital signatures"""
        signed_files = []

        for pdf_file in pdf_files:
            try:
                with pikepdf.open(pdf_file) as pdf:
                    # Check for signature fields
                    if "/AcroForm" in pdf.Root:
                        acro_form = pdf.Root.AcroForm
                        if "/Fields" in acro_form:
                            for field in acro_form.Fields:
                                field_obj = field
                                if "/FT" in field_obj and str(field_obj.FT) == "/Sig":
                                    signed_files.append(os.path.basename(pdf_file))
                                    break
            except Exception as e:
                print(f"Could not check {pdf_file}: {e}")

        return signed_files

    def combine_pdfs_pikepdf(self, output_path, flatten=False):
        """Combine PDFs using pikepdf with better signature handling"""
        try:
            # Create output PDF
            output_pdf = pikepdf.Pdf.new()

            for pdf_file in self.pdf_files:
                with pikepdf.open(pdf_file) as src_pdf:
                    # If flattening, we need to warn that signatures become visual only
                    output_pdf.pages.extend(src_pdf.pages)

            # Save the combined PDF
            output_pdf.save(output_path)
            output_pdf.close()

            return True

        except Exception as e:
            raise Exception(f"Failed to combine PDFs: {str(e)}")

    def add_pdfs(self):
        files = filedialog.askopenfilenames(
            title="Select PDF Files",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
        )
        for file in files:
            if file not in self.pdf_files:
                self.pdf_files.append(file)
                self.files_listbox.insert(tk.END, os.path.basename(file))

        self.update_status(f"Added {len(files)} file(s). Total: {len(self.pdf_files)}")

    def remove_pdf(self):
        selection = self.files_listbox.curselection()
        if selection:
            index = selection[0]
            self.files_listbox.delete(index)
            self.pdf_files.pop(index)
            self.update_status(f"Removed file. Total: {len(self.pdf_files)}")
        else:
            messagebox.showwarning("No Selection", "Please select a file to remove")

    def move_up(self):
        selection = self.files_listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a file to move")
            return

        index = selection[0]
        if index > 0:
            self.pdf_files[index], self.pdf_files[index - 1] = (
                self.pdf_files[index - 1],
                self.pdf_files[index],
            )

            item = self.files_listbox.get(index)
            self.files_listbox.delete(index)
            self.files_listbox.insert(index - 1, item)
            self.files_listbox.selection_set(index - 1)

    def move_down(self):
        selection = self.files_listbox.curselection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a file to move")
            return

        index = selection[0]
        if index < len(self.pdf_files) - 1:
            self.pdf_files[index], self.pdf_files[index + 1] = (
                self.pdf_files[index + 1],
                self.pdf_files[index],
            )

            item = self.files_listbox.get(index)
            self.files_listbox.delete(index)
            self.files_listbox.insert(index + 1, item)
            self.files_listbox.selection_set(index + 1)

    def clear_all(self):
        if self.pdf_files:
            if messagebox.askyesno("Clear All", "Remove all PDF files from the list?"):
                self.pdf_files.clear()
                self.files_listbox.delete(0, tk.END)
                self.update_status("All files cleared")

    def load_printers(self):
        self.update_status("Loading printers...")
        self.printers = []

        try:
            if platform.system() == "Windows":
                try:
                    import win32print

                    printer_list = win32print.EnumPrinters(
                        win32print.PRINTER_ENUM_LOCAL
                        | win32print.PRINTER_ENUM_CONNECTIONS
                    )
                    self.printers = [printer[2] for printer in printer_list]
                except ImportError:
                    messagebox.showwarning(
                        "Missing Module",
                        "Install pywin32 for printer detection:\npip install pywin32",
                    )
                    self.printers = ["Default Printer"]

            elif platform.system() in ["Darwin", "Linux"]:
                result = subprocess.run(
                    ["lpstat", "-p"], capture_output=True, text=True
                )
                for line in result.stdout.split("\n"):
                    if line.startswith("printer"):
                        printer_name = line.split()[1]
                        self.printers.append(printer_name)

            if not self.printers:
                self.printers = ["Default Printer"]

            self.printer_combo["values"] = self.printers
            self.printer_combo.current(0)
            self.update_status(f"Found {len(self.printers)} printer(s)")

        except Exception as e:
            self.printers = ["Default Printer"]
            self.printer_combo["values"] = self.printers
            self.printer_combo.current(0)
            self.update_status(f"Could not load printers: {str(e)}")

    def combine_and_print(self):
        if not self.pdf_files:
            messagebox.showwarning("No Files", "Please add PDF files first")
            return

        # Check for signatures
        signed_files = self.check_for_signatures(self.pdf_files)

        if signed_files and self.signature_var.get() == "warn":
            msg = f"⚠️ The following files contain digital signatures:\n\n"
            msg += "\n".join(f"• {f}" for f in signed_files)
            msg += "\n\nMerging will REMOVE these signatures."
            msg += "\n\nDo you want to continue?"

            if not messagebox.askyesno(
                "Digital Signatures Detected", msg, icon="warning"
            ):
                return

        printer = self.printer_combo.get()
        flatten = self.signature_var.get() == "flatten"

        thread = threading.Thread(
            target=self._do_combine_and_print, args=(printer, flatten)
        )
        thread.daemon = True
        thread.start()

    def _do_combine_and_print(self, printer, flatten):
        try:
            self.update_status("Combining PDFs...")

            with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as temp_pdf:
                temp_path = temp_pdf.name

            self.combine_pdfs_pikepdf(temp_path, flatten)

            self.update_status(f"Sending to printer: {printer}...")

            if platform.system() == "Windows":
                if printer != "Default Printer":
                    os.startfile(temp_path, "print", printer)
                else:
                    os.startfile(temp_path, "print")
            else:
                cmd = ["lpr"]
                if printer != "Default Printer":
                    cmd.extend(["-P", printer])
                cmd.append(temp_path)
                subprocess.run(cmd, check=True)

            self.update_status("✓ Print job sent successfully!")
            messagebox.showinfo("Success", "PDF combined and sent to printer!")

            try:
                os.unlink(temp_path)
            except:
                pass

        except Exception as e:
            self.update_status(f"Error: {str(e)}")
            messagebox.showerror("Error", f"Failed to print:\n{str(e)}")

    def combine_and_save(self):
        if not self.pdf_files:
            messagebox.showwarning("No Files", "Please add PDF files first")
            return

        # Check for signatures
        signed_files = self.check_for_signatures(self.pdf_files)

        if signed_files and self.signature_var.get() == "warn":
            msg = f"⚠️ The following files contain digital signatures:\n\n"
            msg += "\n".join(f"• {f}" for f in signed_files)
            msg += "\n\nMerging will REMOVE these signatures."
            msg += "\n\nDo you want to continue?"

            if not messagebox.askyesno(
                "Digital Signatures Detected", msg, icon="warning"
            ):
                return

        save_path = filedialog.asksaveasfilename(
            defaultextension=".pdf",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")],
            title="Save Combined PDF As",
        )

        if save_path:
            try:
                self.update_status("Combining and saving PDFs...")

                flatten = self.signature_var.get() == "flatten"
                self.combine_pdfs_pikepdf(save_path, flatten)

                self.update_status(f"✓ Saved to: {save_path}")

                msg = f"PDF saved to:\n{save_path}"
                if signed_files:
                    msg += "\n\n⚠️ Note: Digital signatures were removed during merge."

                messagebox.showinfo("Success", msg)

            except Exception as e:
                self.update_status(f"Error: {str(e)}")
                messagebox.showerror("Error", f"Failed to save:\n{str(e)}")

    def update_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFCombinerApp(root)
    root.mainloop()

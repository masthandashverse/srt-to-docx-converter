"""
SRT to DOCX Converter - Main Application

A desktop GUI application that batch converts SRT subtitle files
into formatted Word documents (DOCX).

Usage:
    python converter.py

Requirements:
    pip install python-docx
"""

import os
import sys
import threading
import platform
import subprocess
from pathlib import Path

# ─── Auto-install python-docx if missing ─────────────────────
try:
    from docx import Document
except ImportError:
    print("Installing required package: python-docx ...")
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", "python-docx"]
    )
    print("Installation complete. Restarting...")
    from docx import Document

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from srt_parser import SRTParser
from docx_writer import DOCXWriter


# ═══════════════════════════════════════════════════════════════
# MAIN APPLICATION CLASS
# ═══════════════════════════════════════════════════════════════
class SRTtoDocxApp:
    """
    Main GUI application for converting SRT files to DOCX documents.

    Features:
    - Browse and scan source folder for SRT files
    - Select output folder for converted files
    - Multiple output format styles
    - Progress tracking with real-time updates
    - Threaded conversion (non-blocking GUI)
    - Cross-platform support
    """

    # Application constants
    APP_TITLE = "SRT to DOCX Converter"
    APP_WIDTH = 900
    APP_HEIGHT = 750
    BG_COLOR = '#f0f4f8'
    ACCENT_COLOR = '#1a478a'

    def __init__(self):
        """Initialize the application."""
        self.root = tk.Tk()
        self.root.title(self.APP_TITLE)
        self.root.geometry(f"{self.APP_WIDTH}x{self.APP_HEIGHT}")
        self.root.resizable(True, True)
        self.root.minsize(700, 600)
        self.root.configure(bg=self.BG_COLOR)

        # Set application icon (cross-platform safe)
        try:
            if platform.system() == 'Windows':
                self.root.iconbitmap(default='')
        except tk.TclError:
            pass

        # Core components
        self.parser = SRTParser()
        self.writer = DOCXWriter()

        # State variables
        self.source_folder = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.format_var = tk.StringVar(value="table")
        self.recursive_var = tk.BooleanVar(value=True)
        self.srt_files = []
        self.is_converting = False

        # Build the interface
        self._setup_styles()
        self._build_gui()
        self._center_window()

    # ─────────────────────────────────────────────
    # STYLE CONFIGURATION
    # ─────────────────────────────────────────────
    def _setup_styles(self):
        """Configure ttk widget styles."""
        self.style = ttk.Style()
        self.style.theme_use('clam')

        # Primary button (Convert)
        self.style.configure(
            'Primary.TButton',
            font=('Segoe UI', 11, 'bold'),
            padding=(20, 12),
            background='#1a478a',
            foreground='white'
        )
        self.style.map('Primary.TButton', background=[
            ('active', '#0d2f5e'),
            ('disabled', '#b0b0b0')
        ])

        # Secondary button (Browse, Scan)
        self.style.configure(
            'Secondary.TButton',
            font=('Segoe UI', 9),
            padding=(15, 8),
            background='#5a9bd5',
            foreground='white'
        )
        self.style.map('Secondary.TButton', background=[
            ('active', '#3d7ab5'),
            ('disabled', '#cccccc')
        ])

        # Danger button (Clear)
        self.style.configure(
            'Danger.TButton',
            font=('Segoe UI', 9),
            padding=(15, 8),
            background='#e74c3c',
            foreground='white'
        )
        self.style.map('Danger.TButton', background=[
            ('active', '#c0392b'),
            ('disabled', '#cccccc')
        ])

        # Labels
        self.style.configure('TLabel', font=('Segoe UI', 10), background=self.BG_COLOR)
        self.style.configure(
            'Header.TLabel',
            font=('Segoe UI', 24, 'bold'),
            background=self.BG_COLOR,
            foreground=self.ACCENT_COLOR
        )
        self.style.configure(
            'SubHeader.TLabel',
            font=('Segoe UI', 10),
            background=self.BG_COLOR,
            foreground='#666666'
        )
        self.style.configure(
            'Status.TLabel',
            font=('Segoe UI', 9),
            background=self.BG_COLOR,
            foreground='#555555'
        )

        # Label frames
        self.style.configure('TLabelframe', background=self.BG_COLOR)
        self.style.configure(
            'TLabelframe.Label',
            font=('Segoe UI', 10, 'bold'),
            background=self.BG_COLOR,
            foreground=self.ACCENT_COLOR
        )

        # Checkbutton
        self.style.configure('TCheckbutton', background=self.BG_COLOR)

    # ─────────────────────────────────────────────
    # GUI CONSTRUCTION
    # ─────────────────────────────────────────────
    def _build_gui(self):
        """Build the complete GUI layout."""

        # Scrollable main container
        main_frame = tk.Frame(self.root, bg=self.BG_COLOR)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)

        # ── Header ──
        self._build_header(main_frame)

        # ── Source Folder ──
        self._build_source_section(main_frame)

        # ── Output Folder ──
        self._build_output_section(main_frame)

        # ── Format Options ──
        self._build_format_section(main_frame)

        # ── File List ──
        self._build_file_list_section(main_frame)

        # ── Progress ──
        self._build_progress_section(main_frame)

        # ── Action Buttons ──
        self._build_action_buttons(main_frame)

    def _build_header(self, parent):
        """Build the header section."""
        header_frame = tk.Frame(parent, bg=self.BG_COLOR)
        header_frame.pack(fill=tk.X, pady=(0, 15))

        ttk.Label(
            header_frame,
            text="📝 SRT to DOCX Converter",
            style='Header.TLabel'
        ).pack()

        ttk.Label(
            header_frame,
            text="Batch convert subtitle files (.srt) to Word documents (.docx)",
            style='SubHeader.TLabel'
        ).pack()

    def _build_source_section(self, parent):
        """Build source folder selection section."""
        frame = ttk.LabelFrame(
            parent,
            text="  📁 Source Folder (SRT Files)  ",
            padding=12
        )
        frame.pack(fill=tk.X, pady=(0, 8))

        inner = tk.Frame(frame, bg=self.BG_COLOR)
        inner.pack(fill=tk.X)

        self.source_entry = ttk.Entry(
            inner,
            textvariable=self.source_folder,
            font=('Segoe UI', 10)
        )
        self.source_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        ttk.Button(
            inner,
            text="📂 Browse...",
            command=self.browse_source,
            style='Secondary.TButton'
        ).pack(side=tk.RIGHT)

        # Recursive option
        options_frame = tk.Frame(frame, bg=self.BG_COLOR)
        options_frame.pack(fill=tk.X, pady=(8, 0))

        ttk.Checkbutton(
            options_frame,
            text="Search subdirectories recursively",
            variable=self.recursive_var
        ).pack(side=tk.LEFT)

    def _build_output_section(self, parent):
        """Build output folder selection section."""
        frame = ttk.LabelFrame(
            parent,
            text="  💾 Output Folder (Save DOCX Files)  ",
            padding=12
        )
        frame.pack(fill=tk.X, pady=(0, 8))

        inner = tk.Frame(frame, bg=self.BG_COLOR)
        inner.pack(fill=tk.X)

        self.output_entry = ttk.Entry(
            inner,
            textvariable=self.output_folder,
            font=('Segoe UI', 10)
        )
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))

        ttk.Button(
            inner,
            text="📂 Browse...",
            command=self.browse_output,
            style='Secondary.TButton'
        ).pack(side=tk.RIGHT)

    def _build_format_section(self, parent):
        """Build format options section."""
        frame = ttk.LabelFrame(
            parent,
            text="  ⚙️ Output Format Style  ",
            padding=12
        )
        frame.pack(fill=tk.X, pady=(0, 8))

        format_options = [
            ("📊 Table", "table", "Structured table with index, timestamps, and text columns"),
            ("📄 Plain", "plain", "Numbered entries with visual separators between each"),
            ("🎨 Formatted", "formatted", "Clean paragraphs with small inline timestamps"),
            ("📝 Text Only", "text_only", "Pure subtitle text — no numbers or timestamps"),
            ("🎬 Script", "script", "Screenplay / dialogue style centered format"),
        ]

        for label, value, description in format_options:
            row = tk.Frame(frame, bg=self.BG_COLOR)
            row.pack(fill=tk.X, pady=2)

            ttk.Radiobutton(
                row,
                text=label,
                variable=self.format_var,
                value=value
            ).pack(side=tk.LEFT)

            tk.Label(
                row,
                text=f"  — {description}",
                font=('Segoe UI', 8),
                fg='#888888',
                bg=self.BG_COLOR
            ).pack(side=tk.LEFT, padx=(5, 0))

    def _build_file_list_section(self, parent):
        """Build file list display section."""
        frame = ttk.LabelFrame(
            parent,
            text="  📋 SRT Files Found  ",
            padding=10
        )
        frame.pack(fill=tk.BOTH, expand=True, pady=(0, 8))

        # Listbox with scrollbar
        list_container = tk.Frame(frame, bg=self.BG_COLOR)
        list_container.pack(fill=tk.BOTH, expand=True)

        y_scrollbar = ttk.Scrollbar(list_container, orient=tk.VERTICAL)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        x_scrollbar = ttk.Scrollbar(list_container, orient=tk.HORIZONTAL)
        x_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)

        self.file_listbox = tk.Listbox(
            list_container,
            font=('Consolas', 10),
            selectmode=tk.EXTENDED,
            bg='white',
            fg='#333333',
            selectbackground=self.ACCENT_COLOR,
            selectforeground='white',
            activestyle='none',
            bd=1,
            relief=tk.SOLID,
            yscrollcommand=y_scrollbar.set,
            xscrollcommand=x_scrollbar.set
        )
        self.file_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        y_scrollbar.config(command=self.file_listbox.yview)
        x_scrollbar.config(command=self.file_listbox.xview)

        # File count label
        self.file_count_label = ttk.Label(
            frame,
            text="No files loaded. Select a source folder and click 'Scan'.",
            style='SubHeader.TLabel'
        )
        self.file_count_label.pack(anchor=tk.W, pady=(5, 0))

    def _build_progress_section(self, parent):
        """Build progress bar and status section."""
        frame = tk.Frame(parent, bg=self.BG_COLOR)
        frame.pack(fill=tk.X, pady=(0, 8))

        self.progress_var = tk.DoubleVar(value=0)
        self.progress_bar = ttk.Progressbar(
            frame,
            variable=self.progress_var,
            maximum=100,
            mode='determinate',
            length=500
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 4))

        self.status_label = ttk.Label(
            frame,
            text="Ready — Waiting for input...",
            style='Status.TLabel'
        )
        self.status_label.pack(fill=tk.X)

    def _build_action_buttons(self, parent):
        """Build action button row."""
        frame = tk.Frame(parent, bg=self.BG_COLOR)
        frame.pack(fill=tk.X, pady=(0, 5))

        # Left side buttons
        left_buttons = tk.Frame(frame, bg=self.BG_COLOR)
        left_buttons.pack(side=tk.LEFT)

        self.scan_btn = ttk.Button(
            left_buttons,
            text="🔍 Scan for SRT Files",
            command=self.scan_files,
            style='Secondary.TButton'
        )
        self.scan_btn.pack(side=tk.LEFT, padx=(0, 8))

        self.clear_btn = ttk.Button(
            left_buttons,
            text="🗑️ Clear All",
            command=self.clear_all,
            style='Danger.TButton'
        )
        self.clear_btn.pack(side=tk.LEFT)

        # Right side button
        self.convert_btn = ttk.Button(
            frame,
            text="🚀 Convert All Files",
            command=self.start_conversion,
            style='Primary.TButton'
        )
        self.convert_btn.pack(side=tk.RIGHT)

    # ─────────────────────────────────────────────
    # USER ACTIONS
    # ─────────────────────────────────────────────
    def browse_source(self):
        """Open dialog to select source folder."""
        folder = filedialog.askdirectory(
            title="Select Folder Containing SRT Files",
            mustexist=True
        )
        if folder:
            self.source_folder.set(folder)
            self.scan_files()

    def browse_output(self):
        """Open dialog to select output folder."""
        folder = filedialog.askdirectory(
            title="Select Folder to Save DOCX Files"
        )
        if folder:
            self.output_folder.set(folder)

    def scan_files(self):
        """Scan the source folder for SRT files."""
        source = self.source_folder.get().strip()

        if not source:
            messagebox.showwarning(
                "No Folder Selected",
                "Please select a source folder first!"
            )
            return

        if not os.path.isdir(source):
            messagebox.showerror(
                "Invalid Folder",
                f"The selected folder does not exist:\n{source}"
            )
            return

        # Clear previous results
        self.file_listbox.delete(0, tk.END)
        self.srt_files = []
        self.progress_var.set(0)

        self._update_status("🔍 Scanning for SRT files...")

        # Find SRT files
        recursive = self.recursive_var.get()
        self.srt_files = SRTParser.find_srt_files(source, recursive=recursive)

        if not self.srt_files:
            self.file_count_label.config(
                text="⚠️ No SRT files found in the selected folder!"
            )
            self._update_status("No SRT files found.")
            messagebox.showinfo(
                "No Files Found",
                "No .srt files were found in the selected folder.\n\n"
                "Make sure the folder contains subtitle files with .srt extension."
            )
            return

        # Populate listbox
        total_size = 0
        for filepath in self.srt_files:
            file_info = self.parser.get_file_info(filepath)
            rel_path = os.path.relpath(filepath, source)
            display_text = f"  📄 {rel_path}    ({file_info['size_formatted']})"
            self.file_listbox.insert(tk.END, display_text)
            total_size += file_info['size_bytes']

        count = len(self.srt_files)
        total_size_str = SRTParser._format_size(total_size)
        self.file_count_label.config(
            text=f"✅ Found {count} SRT file{'s' if count != 1 else ''}  "
                 f"(Total: {total_size_str})"
        )
        self._update_status(
            f"Found {count} SRT file(s). Ready to convert."
        )

    def start_conversion(self):
        """Validate inputs and start conversion in background thread."""
        if self.is_converting:
            messagebox.showwarning(
                "In Progress",
                "A conversion is already in progress. Please wait."
            )
            return

        if not self.srt_files:
            messagebox.showwarning(
                "No Files",
                "No SRT files to convert!\n\n"
                "Please select a source folder and scan for files first."
            )
            return

        # Ensure output folder is set
        output = self.output_folder.get().strip()
        if not output:
            output = filedialog.askdirectory(
                title="Select Folder to Save DOCX Files"
            )
            if not output:
                return
            self.output_folder.set(output)

        # Create output folder if needed
        try:
            os.makedirs(output, exist_ok=True)
        except OSError as e:
            messagebox.showerror(
                "Folder Error",
                f"Could not create output folder:\n{output}\n\nError: {e}"
            )
            return

        # Confirmation dialog
        count = len(self.srt_files)
        style = self.format_var.get()
        confirmed = messagebox.askyesno(
            "Confirm Conversion",
            f"Convert {count} SRT file{'s' if count != 1 else ''}?\n\n"
            f"Format: {style.replace('_', ' ').title()}\n"
            f"Output: {output}\n\n"
            f"Proceed?"
        )

        if not confirmed:
            return

        # Disable UI and start conversion
        self.is_converting = True
        self._set_buttons_state(False)
        self.progress_var.set(0)

        thread = threading.Thread(target=self._convert_files, daemon=True)
        thread.start()

    def clear_all(self):
        """Reset the application to its initial state."""
        if self.is_converting:
            messagebox.showwarning(
                "In Progress",
                "Cannot clear while conversion is in progress."
            )
            return

        self.source_folder.set("")
        self.output_folder.set("")
        self.file_listbox.delete(0, tk.END)
        self.srt_files = []
        self.progress_var.set(0)
        self.file_count_label.config(
            text="No files loaded. Select a source folder and click 'Scan'."
        )
        self._update_status("Ready — Waiting for input...")

    # ─────────────────────────────────────────────
    # CONVERSION LOGIC (BACKGROUND THREAD)
    # ─────────────────────────────────────────────
    def _convert_files(self):
        """
        Convert all SRT files to DOCX.
        Runs in a background thread to keep GUI responsive.
        """
        output_root = self.output_folder.get().strip()
        source_root = self.source_folder.get().strip()
        format_style = self.format_var.get()

        total = len(self.srt_files)
        successful = 0
        failed = 0
        errors = []

        for i, srt_path in enumerate(self.srt_files):
            filename = os.path.basename(srt_path)

            # Update status on GUI thread
            self.root.after(0, self._update_status,
                            f"⏳ Converting ({i + 1}/{total}): {filename}")

            try:
                # Parse SRT file
                subtitles = self.parser.parse_file(srt_path)

                if not subtitles:
                    errors.append(
                        f"'{filename}': No valid subtitles found in file."
                    )
                    failed += 1
                    continue

                # Build output path (preserve subfolder structure)
                rel_path = os.path.relpath(srt_path, source_root)
                rel_dir = os.path.dirname(rel_path)
                output_dir = os.path.join(output_root, rel_dir) if rel_dir else output_root
                os.makedirs(output_dir, exist_ok=True)

                # Generate output filename
                base_name = os.path.splitext(filename)[0]
                output_path = os.path.join(output_dir, f"{base_name}.docx")

                # Handle filename conflicts
                counter = 1
                while os.path.exists(output_path):
                    output_path = os.path.join(
                        output_dir, f"{base_name} ({counter}).docx"
                    )
                    counter += 1

                # Create DOCX document
                self.writer.create_document(
                    subtitles=subtitles,
                    source_filename=filename,
                    output_path=output_path,
                    style=format_style
                )

                successful += 1

            except FileNotFoundError as e:
                failed += 1
                errors.append(f"'{filename}': File not found — {e}")
            except ValueError as e:
                failed += 1
                errors.append(f"'{filename}': Parse error — {e}")
            except PermissionError as e:
                failed += 1
                errors.append(f"'{filename}': Permission denied — {e}")
            except Exception as e:
                failed += 1
                errors.append(f"'{filename}': Unexpected error — {type(e).__name__}: {e}")

            # Update progress
            progress = ((i + 1) / total) * 100
            self.root.after(0, self._update_progress, progress)

        # Complete — update UI on main thread
        self.root.after(
            0, self._conversion_complete,
            successful, failed, total, errors
        )

    def _conversion_complete(self, successful, failed, total, errors):
        """Handle conversion completion (runs on main thread)."""
        self.is_converting = False
        self._set_buttons_state(True)
        self.progress_var.set(100)

        # Build result message
        lines = ["═" * 40, "CONVERSION COMPLETE", "═" * 40, ""]
        lines.append(f"✅ Successful:  {successful} / {total}")

        if failed > 0:
            lines.append(f"❌ Failed:      {failed} / {total}")

        lines.append("")
        lines.append(f"📁 Output folder:")
        lines.append(f"   {self.output_folder.get()}")

        if errors:
            lines.append("")
            lines.append("⚠️ Errors:")
            max_display = 15
            for error in errors[:max_display]:
                lines.append(f"   • {error}")
            if len(errors) > max_display:
                lines.append(f"   ... and {len(errors) - max_display} more errors")

        result_message = "\n".join(lines)

        # Update status bar
        if failed > 0:
            status_text = f"⚠️ Done: {successful} converted, {failed} failed out of {total}."
        else:
            status_text = f"✅ Success! All {successful} file(s) converted."

        self._update_status(status_text)

        # Show result dialog
        if failed > 0:
            messagebox.showwarning("Conversion Complete", result_message)
        else:
            messagebox.showinfo("Conversion Complete", result_message)

            # Offer to open output folder
            open_folder = messagebox.askyesno(
                "Open Output Folder",
                "All files converted successfully!\n\n"
                "Would you like to open the output folder?"
            )
            if open_folder:
                self._open_folder(self.output_folder.get())

    # ─────────────────────────────────────────────
    # HELPER METHODS
    # ─────────────────────────────────────────────
    def _update_status(self, message):
        """Update the status label text."""
        self.status_label.config(text=message)

    def _update_progress(self, value):
        """Update the progress bar value."""
        self.progress_var.set(value)

    def _set_buttons_state(self, enabled):
        """Enable or disable all interactive buttons."""
        state = 'normal' if enabled else 'disabled'
        for btn in [self.scan_btn, self.clear_btn, self.convert_btn]:
            btn.configure(state=state)

    def _center_window(self):
        """Center the application window on screen."""
        self.root.update_idletasks()
        screen_w = self.root.winfo_screenwidth()
        screen_h = self.root.winfo_screenheight()
        x = (screen_w // 2) - (self.APP_WIDTH // 2)
        y = (screen_h // 2) - (self.APP_HEIGHT // 2)
        self.root.geometry(f"{self.APP_WIDTH}x{self.APP_HEIGHT}+{x}+{y}")

    @staticmethod
    def _open_folder(path):
        """Open a folder in the system's file explorer."""
        system = platform.system()
        try:
            if system == 'Windows':
                os.startfile(path)
            elif system == 'Darwin':
                subprocess.Popen(['open', path])
            else:
                subprocess.Popen(['xdg-open', path])
        except Exception as e:
            messagebox.showwarning(
                "Could Not Open Folder",
                f"Failed to open folder:\n{path}\n\nError: {e}"
            )

    # ─────────────────────────────────────────────
    # RUN
    # ─────────────────────────────────────────────
    def run(self):
        """Start the application main loop."""
        self.root.mainloop()


# ═══════════════════════════════════════════════════════════════
# COMMAND LINE INTERFACE (FALLBACK)
# ═══════════════════════════════════════════════════════════════
def run_cli():
    """
    Command-line fallback for environments without display.
    """
    print("=" * 50)
    print("  📝 SRT to DOCX Converter (CLI Mode)")
    print("=" * 50)
    print()

    # Get source folder
    source = input("Enter source folder path (containing .srt files): ").strip()
    if not source or not os.path.isdir(source):
        print(f"❌ Error: Invalid folder path: '{source}'")
        sys.exit(1)

    # Get output folder
    output = input("Enter output folder path (to save .docx files): ").strip()
    if not output:
        print("❌ Error: No output folder specified.")
        sys.exit(1)

    os.makedirs(output, exist_ok=True)

    # Choose format
    print("\nAvailable formats:")
    formats = ['table', 'plain', 'formatted', 'text_only', 'script']
    for i, fmt in enumerate(formats, 1):
        print(f"  {i}. {fmt}")

    choice = input(f"\nSelect format (1-{len(formats)}) [default: 1]: ").strip()
    try:
        idx = int(choice) - 1 if choice else 0
        style = formats[idx]
    except (ValueError, IndexError):
        style = 'table'

    print(f"\n🔍 Scanning for SRT files in: {source}")
    parser = SRTParser()
    writer = DOCXWriter()

    srt_files = SRTParser.find_srt_files(source, recursive=True)

    if not srt_files:
        print("❌ No .srt files found!")
        sys.exit(1)

    print(f"✅ Found {len(srt_files)} SRT file(s)\n")

    successful = 0
    failed = 0

    for i, srt_path in enumerate(srt_files, 1):
        filename = os.path.basename(srt_path)
        print(f"  [{i}/{len(srt_files)}] Converting: {filename} ... ", end="")

        try:
            subtitles = parser.parse_file(srt_path)
            if not subtitles:
                print("⚠️ SKIPPED (no subtitles)")
                failed += 1
                continue

            base_name = os.path.splitext(filename)[0]
            output_path = os.path.join(output, f"{base_name}.docx")

            counter = 1
            while os.path.exists(output_path):
                output_path = os.path.join(output, f"{base_name} ({counter}).docx")
                counter += 1

            writer.create_document(subtitles, filename, output_path, style=style)
            print("✅ OK")
            successful += 1

        except Exception as e:
            print(f"❌ ERROR: {e}")
            failed += 1

    print(f"\n{'=' * 50}")
    print(f"  ✅ Successful: {successful}")
    print(f"  ❌ Failed:     {failed}")
    print(f"  📁 Output:     {output}")
    print(f"{'=' * 50}")


# ═══════════════════════════════════════════════════════════════
# ENTRY POINT
# ═══════════════════════════════════════════════════════════════
def main():
    """Application entry point."""
    # Check for --cli flag
    if '--cli' in sys.argv:
        run_cli()
        return

    # Try GUI mode
    try:
        app = SRTtoDocxApp()
        app.run()
    except tk.TclError:
        print("⚠️ No display found. Falling back to CLI mode.\n")
        run_cli()


if __name__ == "__main__":
    main()
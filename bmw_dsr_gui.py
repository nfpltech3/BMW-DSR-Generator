import tkinter as tk
from tkinter import filedialog, messagebox
import os
from datetime import datetime
import os.path
import sys

from bmw_dsr_processor import generate_bmw_dsr

# --- Nagarkot Brand Colors ---
COLOR_PRIMARY_BLUE = "#1F3F6E"
COLOR_ACCENT_RED = "#D8232A"
COLOR_DARK_TEXT = "#1E1E1E"
COLOR_MUTED_GRAY = "#6B7280"
COLOR_LIGHT_BG = "#F4F6F8"
COLOR_PANEL_WHITE = "#FFFFFF"
COLOR_BORDER_GRAY = "#E5E7EB"
COLOR_HOVER_BLUE = "#2A528F"

class BMWDSRApp:
    def __init__(self, root):
        self.root = root
        self.root.title("BMW DSR Generator – Docs Team")
        self.root.configure(bg=COLOR_LIGHT_BG)
        
        # Convert root window to full-screen (maximized) per protocol
        try:
            self.root.state('zoomed')
        except:
            self.root.geometry("1024x768") # Fallback if zoomed fails

        self.input_file = tk.StringVar()

        self.build_ui()

    def build_ui(self):
        # ---------------------------------------------------------
        # HEADER (Dynamic Height)
        # ---------------------------------------------------------
        header_frame = tk.Frame(self.root, bg=COLOR_PANEL_WHITE, height=60)
        header_frame.pack(side=tk.TOP, fill=tk.X)
        header_frame.pack_propagate(False) # Force keeping height
        
        # Bottom border for header
        tk.Frame(self.root, bg=COLOR_BORDER_GRAY, height=1).pack(side=tk.TOP, fill=tk.X)

        # 1. Logo (Left, Fixed Height = 20)
        # Check if logo exists, else text placeholder
        base_path = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
        logo_path = os.path.join(base_path, "logo.png")
        if os.path.exists(logo_path):
            try:
                from PIL import Image, ImageTk
                img = Image.open(logo_path)
                # scale to height 20
                w_percent = (20 / float(img.size[1]))
                h_size = int((float(img.size[0]) * float(w_percent)))
                img = img.resize((h_size, 20), Image.Resampling.LANCZOS)
                self.logo_img = ImageTk.PhotoImage(img)
                logo_label = tk.Label(header_frame, image=self.logo_img, bg=COLOR_PANEL_WHITE)
                logo_label.pack(side=tk.LEFT, padx=20, pady=20)
            except ImportError:
                # PIL not found, fallback to text
                tk.Label(header_frame, text="NAGARKOT LOGO", fg=COLOR_PRIMARY_BLUE, bg=COLOR_PANEL_WHITE, font=("Segoe UI", 12, "bold")).pack(side=tk.LEFT, padx=20, pady=20)
        else:
            # Text placeholder
            tk.Label(header_frame, text="NAGARKOT", fg=COLOR_PRIMARY_BLUE, bg=COLOR_PANEL_WHITE, font=("Segoe UI", 12, "bold", "italic")).pack(side=tk.LEFT, padx=20, pady=20)

        # 2. Centered Title Block (Absolute Center)
        # Using place with relx=0.5 and anchor=tk.CENTER guarantees across-window centering
        title_label = tk.Label(
            header_frame,
            text="BMW DSR Generator",
            fg=COLOR_PRIMARY_BLUE,
            bg=COLOR_PANEL_WHITE,
            font=("Segoe UI", 16, "bold")
        )
        title_label.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
        
        # ---------------------------------------------------------
        # FOOTER (Once per screen, Bottom-Left)
        # ---------------------------------------------------------
        footer_frame = tk.Frame(self.root, bg=COLOR_LIGHT_BG)
        footer_frame.pack(side=tk.BOTTOM, fill=tk.X)
        
        tk.Label(
            footer_frame,
            text="Nagarkot Forwarders Pvt. Ltd. ©",
            fg=COLOR_MUTED_GRAY,
            bg=COLOR_LIGHT_BG,
            font=("Segoe UI", 9)
        ).pack(side=tk.LEFT, padx=20, pady=10)

        # ---------------------------------------------------------
        # BODY (Flexible Content Area)
        # ---------------------------------------------------------
        # Body wrapper for centering using pack wrapper
        body_wrapper = tk.Frame(self.root, bg=COLOR_LIGHT_BG)
        body_wrapper.pack(expand=True, fill=tk.BOTH)

        # Main Panel
        panel = tk.Frame(body_wrapper, bg=COLOR_PANEL_WHITE, padx=40, pady=40, relief=tk.FLAT)
        panel.place(relx=0.5, rely=0.45, anchor=tk.CENTER) # Slightly above dead-center
        
        # Panel Border (simulated with a colored frame wrapping it)
        panel.configure(highlightbackground=COLOR_BORDER_GRAY, highlightcolor=COLOR_BORDER_GRAY, highlightthickness=1)

        # Input Section
        tk.Label(
            panel,
            text="Select Logisys Excel file",
            fg=COLOR_DARK_TEXT,
            bg=COLOR_PANEL_WHITE,
            font=("Segoe UI", 11)
        ).pack(anchor=tk.W, pady=(0, 10))

        input_row = tk.Frame(panel, bg=COLOR_PANEL_WHITE)
        input_row.pack(fill=tk.X, pady=(0, 20))

        entry = tk.Entry(
            input_row,
            textvariable=self.input_file,
            width=50,
            state="readonly",
            font=("Segoe UI", 11),
            fg=COLOR_DARK_TEXT,
            bg=COLOR_LIGHT_BG,
            readonlybackground=COLOR_LIGHT_BG,
            relief=tk.SOLID,
            borderwidth=1,
            highlightbackground=COLOR_BORDER_GRAY
        )
        entry.pack(side=tk.LEFT, ipady=6, padx=(0, 10))

        # Secondary Button (Browse)
        browse_btn = tk.Button(
            input_row,
            text="Browse",
            fg=COLOR_PRIMARY_BLUE,
            bg=COLOR_PANEL_WHITE,
            activebackground=COLOR_LIGHT_BG,
            activeforeground=COLOR_PRIMARY_BLUE,
            font=("Segoe UI", 10, "bold"),
            relief=tk.SOLID,
            borderwidth=1,
            cursor="hand2",
            command=self.browse_file,
            padx=15,
            pady=4
        )
        browse_btn.pack(side=tk.LEFT)

        # Primary Button (Generate)
        generate_btn = tk.Button(
            panel,
            text="Generate BMW DSR",
            fg=COLOR_PANEL_WHITE,
            bg=COLOR_PRIMARY_BLUE,
            activebackground=COLOR_HOVER_BLUE,
            activeforeground=COLOR_PANEL_WHITE,
            font=("Segoe UI", 11, "bold"),
            relief=tk.FLAT,
            cursor="hand2",
            width=30,
            pady=8,
            command=self.generate_dsr
        )
        generate_btn.pack(pady=(10, 5))

        tk.Label(
            panel,
            text="Output will be created in the same folder",
            fg=COLOR_MUTED_GRAY,
            bg=COLOR_PANEL_WHITE,
            font=("Segoe UI", 9)
        ).pack()

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Logisys Excel",
            filetypes=[("Excel Files", "*.xlsx")]
        )
        if file_path:
            self.input_file.set(file_path)

    def generate_dsr(self):
        if not self.input_file.get():
            messagebox.showwarning(
                "No file selected",
                "Please select the Logisys Excel file."
            )
            return

        try:
            input_path = self.input_file.get()
            base_dir = os.path.dirname(input_path)

            timestamp = datetime.now().strftime("%Y%m%d_%H%M")
            output_path = os.path.join(
                base_dir,
                f"BMW_DSR_{timestamp}.xlsx"
            )

            # Change cursor to processing
            self.root.config(cursor="wait")
            self.root.update()

            generate_bmw_dsr(input_path, output_path)

            self.root.config(cursor="")

            messagebox.showinfo(
                "Success",
                f"BMW DSR generated successfully.\n\n{output_path}"
            )

            os.startfile(base_dir)

        except Exception as e:
            self.root.config(cursor="")
            messagebox.showerror(
                "Error",
                f"Failed to generate BMW DSR:\n\n{str(e)}"
            )


if __name__ == "__main__":
    root = tk.Tk()
    app = BMWDSRApp(root)
    root.mainloop()

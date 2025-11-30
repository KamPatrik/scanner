"""
Epson V300 Film Scanner Application
A modern, clean interface for scanning analog film
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
from PIL import Image, ImageTk, ImageOps, ImageFilter, ImageEnhance
import os
from datetime import datetime
import threading
import numpy as np
from PIL.ImageChops import invert as pil_invert
import logging
import traceback
import sys

try:
    import twain
    TWAIN_AVAILABLE = True
except ImportError:
    TWAIN_AVAILABLE = False


class FilmScannerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Film Scanner Pro - Epson V300")
        self.root.geometry("1400x900")
        self.root.configure(bg='#2b2b2b')
        
        # Scanner variables
        self.source_manager = None
        self.scanner = None
        self.scanner_name = ""
        self.is_wia = False
        self.is_wia = False
        self.preview_image = None
        self.preview_image_original = None
        self.scanned_images = []
        
        # Image adjustment variables
        self.brightness = tk.DoubleVar(value=1.0)
        self.contrast = tk.DoubleVar(value=1.0)
        self.exposure = tk.DoubleVar(value=0.0)
        self.invert_negative = tk.BooleanVar(value=False)
        self.remove_dust = tk.BooleanVar(value=False)
        self.rotation_angle = tk.IntVar(value=0)
        self.flip_horizontal = tk.BooleanVar(value=False)
        self.flip_vertical = tk.BooleanVar(value=False)
        
        # Crop variables
        self.crop_active = False
        self.crop_start = None
        self.crop_end = None
        self.crop_rect = None
        
        # Queue variables
        self.scan_queue = []
        self.queue_processing = False
        self.queue_paused = False
        
        # Debug mode and logging
        self.debug_mode = tk.BooleanVar(value=False)
        self.setup_logging()
        
        # Settings
        self.resolution = tk.IntVar(value=2400)
        self.color_mode = tk.StringVar(value="Color")
        self.file_format = tk.StringVar(value="TIFF")
        default_output = os.path.normpath(os.path.join(os.path.expanduser("~"), "Desktop", "Scans"))
        self.output_dir = tk.StringVar(value=default_output)
        self.auto_increment = tk.BooleanVar(value=True)
        self.auto_detect = tk.BooleanVar(value=True)
        self.scan_counter = 1
        
        self.setup_ui()
        
        # Initialize scanner after window is shown
        self.root.after(100, self.initialize_scanner)
        
        # Bind adjustment changes to update preview
        self.brightness.trace_add('write', self.update_preview_adjustments)
        self.contrast.trace_add('write', self.update_preview_adjustments)
        self.exposure.trace_add('write', self.update_preview_adjustments)
        self.invert_negative.trace_add('write', self.update_preview_adjustments)
        self.remove_dust.trace_add('write', self.update_preview_adjustments)
    
    def setup_ui(self):
        """Create the user interface"""
        # Modern color scheme
        bg_color = '#2b2b2b'
        fg_color = '#ffffff'
        accent_color = '#0078d4'
        panel_color = '#3c3c3c'
        
        # Style configuration
        style = ttk.Style()
        style.theme_use('clam')
        style.configure('TFrame', background=bg_color)
        style.configure('TLabel', background=bg_color, foreground=fg_color, font=('Segoe UI', 10))
        style.configure('TButton', font=('Segoe UI', 10), padding=6)
        style.configure('Accent.TButton', background=accent_color, foreground='white', font=('Segoe UI', 11, 'bold'))
        
        # Main container
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        
        # Left panel - Controls
        left_panel = tk.Frame(main_frame, bg=panel_color, width=350)
        left_panel.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S), padx=(0, 10))
        left_panel.grid_propagate(False)
        
        # Create scrollable frame for left panel
        canvas = tk.Canvas(left_panel, bg=panel_color, highlightthickness=0)
        scrollbar = ttk.Scrollbar(left_panel, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg=panel_color)
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Title
        title_label = tk.Label(scrollable_frame, text="Film Scanner Pro", font=('Segoe UI', 18, 'bold'),
                               bg=panel_color, fg=fg_color)
        title_label.pack(pady=20)
        
        # Scanner Status
        status_frame = tk.LabelFrame(scrollable_frame, text="Scanner Status", bg=panel_color, fg=fg_color,
                                     font=('Segoe UI', 10, 'bold'), padx=10, pady=10)
        status_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.status_label = tk.Label(status_frame, text="Initializing...", bg=panel_color, fg='#ffa500',
                                     font=('Segoe UI', 9))
        self.status_label.pack()
        
        # Scan Settings
        settings_frame = tk.LabelFrame(scrollable_frame, text="Scan Settings", bg=panel_color, fg=fg_color,
                                       font=('Segoe UI', 10, 'bold'), padx=10, pady=10)
        settings_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Resolution
        tk.Label(settings_frame, text="Resolution (DPI):", bg=panel_color, fg=fg_color).grid(row=0, column=0, sticky=tk.W, pady=5)
        resolution_combo = ttk.Combobox(settings_frame, textvariable=self.resolution, width=15,
                                       values=[1200, 2400, 3200, 4800, 6400])
        resolution_combo.grid(row=0, column=1, pady=5)
        
        # Color Mode
        tk.Label(settings_frame, text="Color Mode:", bg=panel_color, fg=fg_color).grid(row=1, column=0, sticky=tk.W, pady=5)
        color_combo = ttk.Combobox(settings_frame, textvariable=self.color_mode, width=15,
                                   values=["Color", "Grayscale", "Black & White"])
        color_combo.grid(row=1, column=1, pady=5)
        
        # File Format
        tk.Label(settings_frame, text="File Format:", bg=panel_color, fg=fg_color).grid(row=2, column=0, sticky=tk.W, pady=5)
        format_combo = ttk.Combobox(settings_frame, textvariable=self.file_format, width=15,
                                    values=["TIFF", "PNG", "JPEG"])
        format_combo.grid(row=2, column=1, pady=5)
        
        # Image Adjustments Frame
        adjust_frame = tk.LabelFrame(scrollable_frame, text="Image Adjustments", bg=panel_color, fg=fg_color,
                                     font=('Segoe UI', 10, 'bold'), padx=10, pady=10)
        adjust_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Brightness
        tk.Label(adjust_frame, text="Brightness:", bg=panel_color, fg=fg_color).grid(row=0, column=0, sticky=tk.W, pady=3)
        brightness_scale = tk.Scale(adjust_frame, from_=0.5, to=2.0, resolution=0.1, orient=tk.HORIZONTAL,
                                   variable=self.brightness, bg=panel_color, fg=fg_color, highlightthickness=0,
                                   length=150, troughcolor='#555555')
        brightness_scale.grid(row=0, column=1, pady=3)
        
        # Contrast
        tk.Label(adjust_frame, text="Contrast:", bg=panel_color, fg=fg_color).grid(row=1, column=0, sticky=tk.W, pady=3)
        contrast_scale = tk.Scale(adjust_frame, from_=0.5, to=2.0, resolution=0.1, orient=tk.HORIZONTAL,
                                 variable=self.contrast, bg=panel_color, fg=fg_color, highlightthickness=0,
                                 length=150, troughcolor='#555555')
        contrast_scale.grid(row=1, column=1, pady=3)
        
        # Exposure
        tk.Label(adjust_frame, text="Exposure:", bg=panel_color, fg=fg_color).grid(row=2, column=0, sticky=tk.W, pady=3)
        exposure_scale = tk.Scale(adjust_frame, from_=-1.0, to=1.0, resolution=0.1, orient=tk.HORIZONTAL,
                                 variable=self.exposure, bg=panel_color, fg=fg_color, highlightthickness=0,
                                 length=150, troughcolor='#555555')
        exposure_scale.grid(row=2, column=1, pady=3)
        
        # Negative Inversion
        tk.Checkbutton(adjust_frame, text="Invert Negative", variable=self.invert_negative,
                      bg=panel_color, fg=fg_color, selectcolor=panel_color,
                      activebackground=panel_color, activeforeground=fg_color).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=3)
        
        # Dust Removal
        tk.Checkbutton(adjust_frame, text="Remove Dust/Scratches", variable=self.remove_dust,
                      bg=panel_color, fg=fg_color, selectcolor=panel_color,
                      activebackground=panel_color, activeforeground=fg_color).grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=3)
        
        # Reset button
        reset_btn = tk.Button(adjust_frame, text="Reset All", command=self.reset_adjustments,
                             bg='#555555', fg='white', relief=tk.FLAT, cursor='hand2')
        reset_btn.grid(row=5, column=0, columnspan=2, pady=5)
        
        # Transform Frame
        transform_frame = tk.LabelFrame(scrollable_frame, text="Transform", bg=panel_color, fg=fg_color,
                                       font=('Segoe UI', 10, 'bold'), padx=10, pady=10)
        transform_frame.pack(fill=tk.X, padx=10, pady=10)
        
        # Rotation buttons
        rotate_frame = tk.Frame(transform_frame, bg=panel_color)
        rotate_frame.pack(fill=tk.X, pady=5)
        
        tk.Button(rotate_frame, text="⟲ 90°", command=lambda: self.rotate_image(-90),
                 bg='#555555', fg='white', relief=tk.FLAT, cursor='hand2', width=8).pack(side=tk.LEFT, padx=2)
        tk.Button(rotate_frame, text="⟳ 90°", command=lambda: self.rotate_image(90),
                 bg='#555555', fg='white', relief=tk.FLAT, cursor='hand2', width=8).pack(side=tk.LEFT, padx=2)
        tk.Button(rotate_frame, text="180°", command=lambda: self.rotate_image(180),
                 bg='#555555', fg='white', relief=tk.FLAT, cursor='hand2', width=8).pack(side=tk.LEFT, padx=2)
        
        # Flip buttons
        flip_frame = tk.Frame(transform_frame, bg=panel_color)
        flip_frame.pack(fill=tk.X, pady=5)
        
        tk.Button(flip_frame, text="Flip H", command=self.flip_h,
                 bg='#555555', fg='white', relief=tk.FLAT, cursor='hand2', width=12).pack(side=tk.LEFT, padx=2)
        tk.Button(flip_frame, text="Flip V", command=self.flip_v,
                 bg='#555555', fg='white', relief=tk.FLAT, cursor='hand2', width=12).pack(side=tk.LEFT, padx=2)
        
        # Crop button
        self.crop_btn = tk.Button(transform_frame, text="✂ Crop Selection", command=self.toggle_crop_mode,
                                 bg='#555555', fg='white', relief=tk.FLAT, cursor='hand2')
        self.crop_btn.pack(fill=tk.X, pady=5)
        
        # Output Directory
        output_frame = tk.LabelFrame(scrollable_frame, text="Output", bg=panel_color, fg=fg_color,
                                     font=('Segoe UI', 10, 'bold'), padx=10, pady=10)
        output_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Label(output_frame, text="Save to:", bg=panel_color, fg=fg_color).pack(anchor=tk.W)
        
        dir_frame = tk.Frame(output_frame, bg=panel_color)
        dir_frame.pack(fill=tk.X, pady=5)
        
        self.dir_label = tk.Label(dir_frame, text=self.output_dir.get()[:30] + "...", 
                                  bg=panel_color, fg='#aaaaaa', anchor=tk.W)
        self.dir_label.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        browse_btn = ttk.Button(dir_frame, text="Browse", command=self.browse_directory)
        browse_btn.pack(side=tk.RIGHT)
        
        tk.Checkbutton(output_frame, text="Auto-increment filenames", variable=self.auto_increment,
                      bg=panel_color, fg=fg_color, selectcolor=panel_color,
                      activebackground=panel_color, activeforeground=fg_color).pack(anchor=tk.W, pady=5)
        
        tk.Checkbutton(output_frame, text="Auto-detect film frames", variable=self.auto_detect,
                      bg=panel_color, fg=fg_color, selectcolor=panel_color,
                      activebackground=panel_color, activeforeground=fg_color).pack(anchor=tk.W, pady=2)
        
        # Action Buttons
        button_frame = tk.Frame(scrollable_frame, bg=panel_color)
        button_frame.pack(fill=tk.X, padx=10, pady=20)
        
        self.preview_btn = tk.Button(button_frame, text="Preview", command=self.preview_scan,
                                     bg='#555555', fg='white', font=('Segoe UI', 11),
                                     relief=tk.FLAT, cursor='hand2', padx=20, pady=10)
        self.preview_btn.pack(fill=tk.X, pady=5)
        
        self.scan_btn = tk.Button(button_frame, text="Scan", command=self.start_scan,
                                  bg=accent_color, fg='white', font=('Segoe UI', 12, 'bold'),
                                  relief=tk.FLAT, cursor='hand2', padx=20, pady=12)
        self.scan_btn.pack(fill=tk.X, pady=5)
        
        self.batch_btn = tk.Button(button_frame, text="Batch Scan", command=self.batch_scan,
                                   bg='#0d6efd', fg='white', font=('Segoe UI', 11),
                                   relief=tk.FLAT, cursor='hand2', padx=20, pady=10)
        self.batch_btn.pack(fill=tk.X, pady=5)
        
        # Queue Frame
        queue_frame = tk.LabelFrame(scrollable_frame, text="Scan Queue", bg=panel_color, fg=fg_color,
                                    font=('Segoe UI', 10, 'bold'), padx=10, pady=10)
        queue_frame.pack(fill=tk.X, padx=10, pady=10)
        
        queue_btn_frame = tk.Frame(queue_frame, bg=panel_color)
        queue_btn_frame.pack(fill=tk.X, pady=5)
        
        tk.Button(queue_btn_frame, text="Add to Queue", command=self.add_to_queue,
                 bg='#6c757d', fg='white', relief=tk.FLAT, cursor='hand2', width=12).pack(side=tk.LEFT, padx=2)
        tk.Button(queue_btn_frame, text="Clear Queue", command=self.clear_queue,
                 bg='#dc3545', fg='white', relief=tk.FLAT, cursor='hand2', width=12).pack(side=tk.LEFT, padx=2)
        
        self.queue_label = tk.Label(queue_frame, text="Queue: 0 scans", bg=panel_color, fg=fg_color,
                                    font=('Segoe UI', 9))
        self.queue_label.pack(pady=5)
        
        self.process_queue_btn = tk.Button(queue_frame, text="▶ Process Queue", command=self.process_queue,
                                          bg='#28a745', fg='white', font=('Segoe UI', 10, 'bold'),
                                          relief=tk.FLAT, cursor='hand2')
        self.process_queue_btn.pack(fill=tk.X, pady=5)
        
        self.pause_queue_btn = tk.Button(queue_frame, text="⏸ Pause Queue", command=self.toggle_pause_queue,
                                        bg='#ffc107', fg='black', font=('Segoe UI', 9),
                                        relief=tk.FLAT, cursor='hand2', state=tk.DISABLED)
        self.pause_queue_btn.pack(fill=tk.X, pady=2)
        
        # Statistics
        stats_frame = tk.LabelFrame(scrollable_frame, text="Session Info", bg=panel_color, fg=fg_color,
                                    font=('Segoe UI', 10, 'bold'), padx=10, pady=10)
        stats_frame.pack(fill=tk.X, padx=10, pady=10)
        
        self.stats_label = tk.Label(stats_frame, text="Scans completed: 0", bg=panel_color, fg=fg_color)
        self.stats_label.pack()
        
        # Debug Frame
        debug_frame = tk.LabelFrame(scrollable_frame, text="Debug Tools", bg=panel_color, fg=fg_color,
                                    font=('Segoe UI', 10, 'bold'), padx=10, pady=10)
        debug_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Checkbutton(debug_frame, text="Enable Debug Mode", variable=self.debug_mode,
                      bg=panel_color, fg=fg_color, selectcolor=panel_color,
                      activebackground=panel_color, activeforeground=fg_color,
                      command=self.toggle_debug_mode).pack(anchor=tk.W, pady=3)
        
        tk.Button(debug_frame, text="View Error Log", command=self.show_error_log,
                 bg='#555555', fg='white', relief=tk.FLAT, cursor='hand2').pack(fill=tk.X, pady=2)
        
        tk.Button(debug_frame, text="Test Connection", command=self.test_scanner_connection,
                 bg='#555555', fg='white', relief=tk.FLAT, cursor='hand2').pack(fill=tk.X, pady=2)
        
        # Right panel - Preview
        right_panel = tk.Frame(main_frame, bg=panel_color)
        right_panel.grid(row=0, column=1, sticky=(tk.W, tk.E, tk.N, tk.S))
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        preview_header = tk.Frame(right_panel, bg=panel_color)
        preview_header.pack(fill=tk.X, pady=10)
        
        preview_label = tk.Label(preview_header, text="Preview", font=('Segoe UI', 14, 'bold'),
                                bg=panel_color, fg=fg_color)
        preview_label.pack(side=tk.LEFT, padx=20)
        
        self.crop_info_label = tk.Label(preview_header, text="", font=('Segoe UI', 9),
                                        bg=panel_color, fg='#ffa500')
        self.crop_info_label.pack(side=tk.RIGHT, padx=20)
        
        # Preview canvas
        self.preview_canvas = tk.Canvas(right_panel, bg='#1a1a1a', highlightthickness=0)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        self.preview_text = self.preview_canvas.create_text(
            400, 300, text="No preview available\nClick 'Preview' to see scan preview",
            fill='#666666', font=('Segoe UI', 12), justify=tk.CENTER
        )
        
        # Bind crop events
        self.preview_canvas.bind("<Button-1>", self.crop_mouse_down)
        self.preview_canvas.bind("<B1-Motion>", self.crop_mouse_drag)
        self.preview_canvas.bind("<ButtonRelease-1>", self.crop_mouse_up)
    
    def reset_adjustments(self):
        """Reset all image adjustments to default"""
        self.brightness.set(1.0)
        self.contrast.set(1.0)
        self.exposure.set(0.0)
        self.invert_negative.set(False)
        self.remove_dust.set(False)
        self.rotation_angle.set(0)
        self.flip_horizontal.set(False)
        self.flip_vertical.set(False)
        self.update_preview_adjustments()
    
    def setup_logging(self):
        """Setup logging system"""
        log_dir = os.path.normpath(os.path.join(os.path.expanduser("~"), "Desktop", "Scans", "logs"))
        os.makedirs(log_dir, exist_ok=True)
        
        log_file = os.path.join(log_dir, f"scanner_log_{datetime.now().strftime('%Y%m%d')}.log")
        
        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler(sys.stdout)
            ]
        )
        
        self.logger = logging.getLogger(__name__)
        self.logger.info("Film Scanner Pro started")
        self.log_file = log_file
    
    def toggle_debug_mode(self):
        """Toggle debug mode"""
        if self.debug_mode.get():
            self.logger.setLevel(logging.DEBUG)
            self.logger.debug("Debug mode enabled")
            messagebox.showinfo("Debug Mode", "Debug mode enabled\nDetailed logs will be saved")
        else:
            self.logger.setLevel(logging.INFO)
            self.logger.info("Debug mode disabled")
    
    def show_error_log(self):
        """Show error log in a window"""
        try:
            with open(self.log_file, 'r') as f:
                log_content = f.read()
            
            log_window = tk.Toplevel(self.root)
            log_window.title("Error Log")
            log_window.geometry("800x600")
            log_window.configure(bg='#2b2b2b')
            
            text_frame = tk.Frame(log_window, bg='#2b2b2b')
            text_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
            
            scrollbar = tk.Scrollbar(text_frame)
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
            
            text_widget = tk.Text(text_frame, wrap=tk.WORD, bg='#1a1a1a', fg='#00ff00',
                                 font=('Consolas', 9), yscrollcommand=scrollbar.set)
            text_widget.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
            scrollbar.config(command=text_widget.yview)
            
            text_widget.insert(tk.END, log_content)
            text_widget.config(state=tk.DISABLED)
            
            # Add buttons
            btn_frame = tk.Frame(log_window, bg='#2b2b2b')
            btn_frame.pack(fill=tk.X, padx=10, pady=5)
            
            tk.Button(btn_frame, text="Refresh", command=lambda: self.refresh_log(text_widget),
                     bg='#555555', fg='white', relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=5)
            tk.Button(btn_frame, text="Clear Log", command=lambda: self.clear_log(text_widget),
                     bg='#dc3545', fg='white', relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=5)
            tk.Button(btn_frame, text="Open Log File", command=self.open_log_file,
                     bg='#0078d4', fg='white', relief=tk.FLAT, cursor='hand2').pack(side=tk.LEFT, padx=5)
        
        except Exception as e:
            messagebox.showerror("Error", f"Could not open log file:\n{str(e)}")
    
    def refresh_log(self, text_widget):
        """Refresh log content"""
        try:
            with open(self.log_file, 'r') as f:
                log_content = f.read()
            text_widget.config(state=tk.NORMAL)
            text_widget.delete(1.0, tk.END)
            text_widget.insert(tk.END, log_content)
            text_widget.config(state=tk.DISABLED)
            text_widget.see(tk.END)
        except Exception as e:
            messagebox.showerror("Error", f"Could not refresh log:\n{str(e)}")
    
    def clear_log(self, text_widget):
        """Clear log file"""
        if messagebox.askyesno("Clear Log", "Clear all log entries?"):
            try:
                with open(self.log_file, 'w') as f:
                    f.write(f"Log cleared at {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                self.refresh_log(text_widget)
                self.logger.info("Log file cleared by user")
            except Exception as e:
                messagebox.showerror("Error", f"Could not clear log:\n{str(e)}")
    
    def open_log_file(self):
        """Open log file in default text editor"""
        try:
            os.startfile(self.log_file)
        except Exception as e:
            messagebox.showerror("Error", f"Could not open log file:\n{str(e)}")
    
    def test_scanner_connection(self):
        """Test scanner connection and capabilities"""
        if not TWAIN_AVAILABLE:
            messagebox.showwarning("TWAIN Not Available", "Cannot test connection: TWAIN library not installed")
            return
        
        try:
            result = "Scanner Connection Test\n" + "="*40 + "\n\n"
            
            if self.source_manager:
                result += "✓ TWAIN Source Manager: OK\n"
                sources = self.source_manager.GetSourceList()
                result += f"✓ Available scanners: {len(sources)}\n"
                for i, source in enumerate(sources, 1):
                    result += f"  {i}. {source}\n"
            else:
                result += "✗ TWAIN Source Manager: Not initialized\n"
            
            if self.scanner:
                result += "\n✓ Scanner connection: OK\n"
                result += f"  Scanner ready for operation\n"
            else:
                result += "\n✗ Scanner connection: Not connected\n"
            
            result += f"\n✓ Image library (Pillow): OK\n"
            result += f"✓ NumPy (frame detection): OK\n"
            result += f"\nOutput directory: {self.output_dir.get()}\n"
            
            if os.path.exists(self.output_dir.get()):
                result += "✓ Output directory exists\n"
            else:
                result += "✗ Output directory does not exist (will be created)\n"
            
            messagebox.showinfo("Connection Test", result)
            self.logger.info("Scanner connection test performed")
            
        except Exception as e:
            error_msg = f"Connection test failed:\n{str(e)}\n\n{traceback.format_exc()}"
            messagebox.showerror("Test Failed", error_msg)
            self.logger.error(f"Connection test failed: {str(e)}")
    
    def rotate_image(self, angle):
        """Rotate preview image"""
        if not self.preview_image_original:
            return
        
        current = self.rotation_angle.get()
        self.rotation_angle.set((current + angle) % 360)
        self.update_preview_adjustments()
    
    def flip_h(self):
        """Flip image horizontally"""
        self.flip_horizontal.set(not self.flip_horizontal.get())
        self.update_preview_adjustments()
    
    def flip_v(self):
        """Flip image vertically"""
        self.flip_vertical.set(not self.flip_vertical.get())
        self.update_preview_adjustments()
    
    def toggle_crop_mode(self):
        """Toggle crop selection mode"""
        self.crop_active = not self.crop_active
        if self.crop_active:
            self.crop_btn.config(bg='#00ff00', text="✓ Crop Active - Draw Rectangle")
            self.crop_info_label.config(text="Draw rectangle to select crop area")
        else:
            self.crop_btn.config(bg='#555555', text="✂ Crop Selection")
            self.crop_info_label.config(text="")
            if self.crop_rect:
                self.preview_canvas.delete(self.crop_rect)
                self.crop_rect = None
            self.crop_start = None
            self.crop_end = None
    
    def crop_mouse_down(self, event):
        """Handle crop selection start"""
        if not self.crop_active:
            return
        self.crop_start = (event.x, event.y)
        if self.crop_rect:
            self.preview_canvas.delete(self.crop_rect)
    
    def crop_mouse_drag(self, event):
        """Handle crop selection drag"""
        if not self.crop_active or not self.crop_start:
            return
        
        if self.crop_rect:
            self.preview_canvas.delete(self.crop_rect)
        
        x0, y0 = self.crop_start
        self.crop_rect = self.preview_canvas.create_rectangle(
            x0, y0, event.x, event.y, outline='#00ff00', width=2
        )
    
    def crop_mouse_up(self, event):
        """Handle crop selection end"""
        if not self.crop_active or not self.crop_start:
            return
        
        self.crop_end = (event.x, event.y)
        self.apply_crop()
    
    def apply_crop(self):
        """Apply crop to the preview image"""
        if not self.preview_image_original or not self.crop_start or not self.crop_end:
            return
        
        # Get canvas and image dimensions
        canvas_width = self.preview_canvas.winfo_width()
        canvas_height = self.preview_canvas.winfo_height()
        
        # Convert canvas coordinates to image coordinates
        img = self.preview_image_original
        
        # Calculate the displayed image size on canvas
        img_ratio = img.width / img.height
        canvas_ratio = canvas_width / canvas_height
        
        if img_ratio > canvas_ratio:
            display_width = canvas_width - 40
            display_height = int(display_width / img_ratio)
        else:
            display_height = canvas_height - 40
            display_width = int(display_height * img_ratio)
        
        # Calculate offset (image is centered)
        offset_x = (canvas_width - display_width) // 2
        offset_y = (canvas_height - display_height) // 2
        
        # Convert crop coordinates
        x1 = int((self.crop_start[0] - offset_x) * img.width / display_width)
        y1 = int((self.crop_start[1] - offset_y) * img.height / display_height)
        x2 = int((self.crop_end[0] - offset_x) * img.width / display_width)
        y2 = int((self.crop_end[1] - offset_y) * img.height / display_height)
        
        # Ensure coordinates are in bounds
        x1, x2 = max(0, min(x1, x2)), min(img.width, max(x1, x2))
        y1, y2 = max(0, min(y1, y2)), min(img.height, max(y1, y2))
        
        if x2 - x1 > 10 and y2 - y1 > 10:  # Minimum crop size
            self.preview_image_original = img.crop((x1, y1, x2, y2))
            self.update_preview_adjustments()
            self.toggle_crop_mode()  # Exit crop mode
            messagebox.showinfo("Crop Applied", "Crop has been applied to the preview")
    
    def update_preview_adjustments(self, *args):
        """Update preview with current adjustments"""
        if not self.preview_image_original:
            return
        
        img = self.preview_image_original.copy()
        
        # Apply transformations
        if self.rotation_angle.get() != 0:
            img = img.rotate(-self.rotation_angle.get(), expand=True)
        
        if self.flip_horizontal.get():
            img = ImageOps.mirror(img)
        
        if self.flip_vertical.get():
            img = ImageOps.flip(img)
        
        # Apply adjustments
        img = self.apply_adjustments(img)
        
        self.preview_image = img
        self.display_preview(img)
    
    def apply_adjustments(self, image):
        """Apply all image adjustments"""
        try:
            img = image.copy()
            
            # Brightness
            if self.brightness.get() != 1.0:
                enhancer = ImageEnhance.Brightness(img)
                img = enhancer.enhance(self.brightness.get())
            
            # Contrast
            if self.contrast.get() != 1.0:
                enhancer = ImageEnhance.Contrast(img)
                img = enhancer.enhance(self.contrast.get())
            
            # Exposure (using brightness on top of brightness for more range)
            if self.exposure.get() != 0.0:
                # Exposure: apply gamma correction
                exposure_factor = 1.0 + self.exposure.get()
                img_array = np.array(img).astype(np.float32)
                img_array = np.clip(img_array * exposure_factor, 0, 255)
                img = Image.fromarray(img_array.astype(np.uint8))
            
            # Negative inversion
            if self.invert_negative.get():
                if img.mode == 'RGB':
                    img = ImageOps.invert(img)
                elif img.mode == 'L':
                    img = ImageOps.invert(img)
            
            # Dust/scratch removal (despeckle)
            if self.remove_dust.get():
                img = img.filter(ImageFilter.MedianFilter(size=3))
            
            return img
            
        except Exception as e:
            self.logger.error(f"Error applying adjustments: {str(e)}")
            messagebox.showwarning("Adjustment Error", 
                f"Could not apply some adjustments:\n{str(e)}\n\nUsing original image.")
            return image
    
    def initialize_scanner(self):
        """Initialize connection to scanner"""
        if not TWAIN_AVAILABLE:
            self.status_label.config(text="TWAIN not available - Install python-twain", fg='#ff4444')
            self.logger.error("TWAIN library not available")
            messagebox.showwarning("TWAIN Not Available",
                                 "Python TWAIN library not installed.\n\n"
                                 "To use this scanner application, please install:\n"
                                 "pip install pytwain\n\n"
                                 "For now, you can test the interface in demo mode.")
            self.scan_btn.config(state=tk.DISABLED)
            self.preview_btn.config(state=tk.DISABLED)
            self.batch_btn.config(state=tk.DISABLED)
            return
        
        # Try multiple initialization methods
        methods = [
            ("Standard TWAIN DSM", lambda: self._init_standard_twain()),
            ("Legacy TWAIN 32-bit", lambda: self._init_legacy_twain()),
            ("WIA fallback", lambda: self._init_wia_fallback())
        ]
        
        for method_name, method in methods:
            try:
                self.logger.info(f"Trying initialization method: {method_name}")
                if method():
                    return
            except Exception as e:
                self.logger.warning(f"{method_name} failed: {str(e)}")
                continue
        
        # All methods failed
        error_msg = "Could not connect to scanner using any method.\n\n" + \
                   "The scanner is detected but fails to open.\n" + \
                   "This is a known issue with some Epson TWAIN drivers.\n\n" + \
                   "Try this fix:\n" + \
                   "1. Open PowerShell and run:\n" + \
                   "   python -m pip install pywin32\n\n" + \
                   "2. Restart this app\n\n" + \
                   "This will enable WIA support which works better\n" + \
                   "with Epson scanners than TWAIN.\n\n" + \
                   "The app will work in demo mode for now."
        
        self.status_label.config(text="Scanner not available", fg='#ff4444')
        self.logger.error("All scanner initialization methods failed")
        messagebox.showwarning("Scanner Not Available", error_msg)
        self.scan_btn.config(state=tk.DISABLED)
        self.preview_btn.config(state=tk.DISABLED)
        self.batch_btn.config(state=tk.DISABLED)
    
    def _init_standard_twain(self):
        """Try standard TWAIN initialization"""
        self.logger.info("Attempting standard TWAIN initialization...")
        
        # Ensure window is ready
        self.root.update()
        
        # Get proper window handle for TWAIN
        import ctypes
        hwnd = int(self.root.wm_frame(), 16)
        
        self.source_manager = twain.SourceManager(hwnd)
        sources = self.source_manager.GetSourceList()
        
        if not sources:
            raise Exception("No scanners detected")
        
        self.logger.info(f"Found {len(sources)} scanner(s): {sources}")
        
        # Reorder sources: AVOID WIA for film scanning, prefer native TWAIN
        # Native TWAIN drivers support transparency/film settings better
        ordered_sources = []
        for source in sources:
            if 'WIA-' not in source:
                # Prefer V300 over V370 if both present
                if 'V30/V300' in source:
                    ordered_sources.insert(0, source)  # V300 at front
                else:
                    ordered_sources.append(source)
        
        # Add WIA as last resort
        for source in sources:
            if 'WIA-' in source:
                ordered_sources.append(source)
        
        # Try each scanner until one works
        last_error = None
        for scanner_name in ordered_sources:
            try:
                self.logger.info(f"Trying to open: {scanner_name}")
                self.scanner = self.source_manager.OpenSource(scanner_name)
                self.scanner_name = scanner_name  # Store for WIA detection
                self.is_wia = 'WIA-' in scanner_name
                self.status_label.config(text=f"Connected: {scanner_name}", fg='#00ff00')
                self.logger.info(f"Successfully connected to scanner: {scanner_name}")
                return True
            except Exception as e:
                self.logger.warning(f"Failed to open {scanner_name}: {str(e)}")
                last_error = e
                continue
        
        # All scanners failed
        raise Exception(f"Could not open any scanner. Last error: {last_error}")
    
    def _init_legacy_twain(self):
        """Try legacy TWAIN with window handle"""
        self.logger.info("Attempting legacy TWAIN initialization...")
        import ctypes
        
        # Ensure window is visible and focused
        self.root.update()
        self.root.focus_force()
        self.root.after(100)  # Small delay for window to be ready
        self.root.update()
        
        # Get window handle
        hwnd = int(self.root.wm_frame(), 16)  # Convert Tk window to hwnd
        self.logger.info(f"Using window handle: {hwnd}")
        
        self.source_manager = twain.SourceManager(hwnd)
        sources = self.source_manager.GetSourceList()
        
        if not sources:
            raise Exception("No scanners detected")
        
        self.logger.info(f"Found {len(sources)} scanner(s): {sources}")
        
        # Reorder sources: AVOID WIA for film scanning, prefer native TWAIN
        ordered_sources = []
        for source in sources:
            if 'WIA-' not in source:
                if 'V30/V300' in source:
                    ordered_sources.insert(0, source)  # V300 first
                else:
                    ordered_sources.append(source)
        
        # Add WIA as last resort
        for source in sources:
            if 'WIA-' in source:
                ordered_sources.append(source)
        
        # Try each scanner until one works
        last_error = None
        for scanner_name in ordered_sources:
            try:
                self.logger.info(f"Trying to open: {scanner_name}")
                self.scanner = self.source_manager.OpenSource(scanner_name)
                self.scanner_name = scanner_name
                self.is_wia = 'WIA-' in scanner_name
                
                self.status_label.config(text=f"Connected: {scanner_name}", fg='#00ff00')
                self.logger.info(f"Legacy TWAIN connected: {scanner_name}")
                return True
            except Exception as e:
                self.logger.warning(f"Failed to open {scanner_name}: {str(e)}")
                last_error = e
                continue
        
        # All scanners failed
        raise Exception(f"Could not open any scanner. Last error: {last_error}")
    
    def _init_wia_fallback(self):
        """Try WIA as fallback (Windows Image Acquisition)"""
        self.logger.info("Attempting WIA fallback...")
        try:
            import win32com.client
            wia = win32com.client.Dispatch("WIA.DeviceManager")
            devices = wia.DeviceInfos
            
            if devices.Count == 0:
                raise Exception("No WIA devices found")
            
            device_name = devices[1].Properties("Name").Value
            self.logger.info(f"Found WIA device: {device_name}")
            self.status_label.config(text=f"Connected (WIA): {device_name}", fg='#00ff00')
            
            # Store WIA device for later use
            self.wia_device = wia.DeviceInfos[1].Connect()
            self.scanner = None  # Mark as WIA mode
            return True
        except ImportError:
            self.logger.warning("WIA not available (pywin32 not installed)")
            return False
        except Exception as e:
            self.logger.warning(f"WIA initialization failed: {str(e)}")
            return False
    
    def browse_directory(self):
        """Browse for output directory"""
        try:
            initial_dir = self.output_dir.get() if os.path.exists(self.output_dir.get()) else os.path.expanduser("~")
            directory = filedialog.askdirectory(initialdir=initial_dir)
            
            if directory:
                # Validate directory is writable
                test_file = os.path.join(directory, ".scanner_test")
                try:
                    with open(test_file, 'w') as f:
                        f.write("test")
                    os.remove(test_file)
                    
                    self.output_dir.set(directory)
                    self.dir_label.config(text=directory[:30] + "...")
                    self.logger.info(f"Output directory changed to: {directory}")
                    
                except PermissionError:
                    messagebox.showerror("Permission Error", 
                                       f"Cannot write to directory:\n{directory}\n\nPlease choose a different location.")
                    self.logger.error(f"Permission denied for directory: {directory}")
                    
        except Exception as e:
            messagebox.showerror("Error", f"Error selecting directory:\n{str(e)}")
            self.logger.error(f"Directory selection error: {str(e)}")
    
    def preview_scan(self):
        """Show preview of scan"""
        if not TWAIN_AVAILABLE or not self.scanner:
            messagebox.showinfo("Demo Mode", "Preview would show here when scanner is connected")
            return
        
        self.status_label.config(text="Generating preview...", fg='#ffa500')
        
        # WIA scanners need UI on main thread
        if self.is_wia or 'WIA-' in self.scanner_name:
            self._do_preview()
        else:
            threading.Thread(target=self._do_preview, daemon=True).start()
    
    def _do_preview(self):
        """Perform preview scan in background thread"""
        try:
            self.logger.info("Starting preview scan...")
            self.logger.info(f"Scanner: {self.scanner_name}, is_wia: {self.is_wia}")
            
            if not self.scanner:
                raise Exception("Scanner not initialized")
            
            # Set up scanner for preview (lower resolution)
            # Configure for film/transparency scanning
            if not self.is_wia and 'WIA-' not in self.scanner_name:
                try:
                    # Set document source to transparency unit (film holder with backlight)
                    try:
                        # CAP_FEEDERENABLED = False for flatbed/transparency
                        self.scanner.SetCapability(twain.CAP_FEEDERENABLED, twain.TWTY_BOOL, False)
                    except:
                        pass
                    
                    try:
                        # Set to transparency/film mode (this enables backlight)
                        # ICAP_LIGHTPATH = 0x1005, TWLP_TRANSMISSIVE = 1
                        self.scanner.SetCapability(0x1005, twain.TWTY_UINT16, 1)
                    except:
                        self.logger.warning("Could not set transparency mode - may need to set in scanner UI")
                    
                    # Set resolution
                    self.scanner.SetCapability(twain.ICAP_XRESOLUTION, twain.TWTY_FIX32, 150)
                    self.scanner.SetCapability(twain.ICAP_YRESOLUTION, twain.TWTY_FIX32, 150)
                except Exception as e:
                    self.logger.warning(f"Could not set all capabilities for preview: {e}")
            
            # Request scan - Always show UI for film scanning configuration
            self.logger.info("Showing scanner UI for film settings")
            
            # Show instruction for film scanning
            messagebox.showinfo("Film Scanning - IMPORTANT!",
                "The Epson Scan window will open.\n\n"
                "CRITICAL STEPS:\n"
                "1. Look for 'Document Type' or mode selector\n"
                "2. Change from 'Reflective' to 'Film' or 'Transparency'\n"
                "3. Select 'Negative Film' or 'Positive Film'\n"
                "4. Choose film holder type (35mm, etc.)\n"
                "5. Preview and select scan area\n\n"
                "The backlight will turn ON when you select Film mode!")
            
            self.scanner.RequestAcquire(1, 1)  # Always show UI for film scanning
            
            # Get image data
            rv = self.scanner.XferImageNatively()
            
            # Handle both return formats: (handle, more) or just handle
            if isinstance(rv, tuple):
                image_handle = rv[0]
            else:
                image_handle = rv
            
            if not image_handle:
                raise Exception("No image data received from scanner")
            
            # Convert DIB handle to PIL Image using pytwain's method
            import tempfile
            temp_bmp = tempfile.mktemp(suffix='.bmp')
            
            try:
                # Use twain module's DIBToBMFile if available
                import twain
                if hasattr(twain, 'DIBToBMFile'):
                    twain.DIBToBMFile(image_handle, temp_bmp)
                    self.preview_image_original = Image.open(temp_bmp)
                else:
                    # Fallback: save handle as temp file and open
                    # This is what pytwain does internally
                    from PIL import ImageWin
                    self.preview_image_original = ImageWin.Dib(image_handle).image
            finally:
                if os.path.exists(temp_bmp):
                    try:
                        os.remove(temp_bmp)
                    except:
                        pass
            
            if self.preview_image_original.size[0] == 0 or self.preview_image_original.size[1] == 0:
                raise Exception("Invalid image dimensions received")
            
            self.preview_image = self.preview_image_original.copy()
            self.logger.info(f"Preview scan successful: {self.preview_image_original.size}")
            self.update_preview_adjustments()
            
            self.root.after(0, lambda: self.status_label.config(text="Preview ready", fg='#00ff00'))
            
        except Exception as e:
            error_msg = f"Preview failed: {str(e)}"
            self.logger.error(f"{error_msg}\n{traceback.format_exc()}")
            self.root.after(0, lambda: self.status_label.config(text=error_msg[:50], fg='#ff4444'))
            self.root.after(0, lambda: messagebox.showerror("Preview Error", 
                f"{error_msg}\n\nCheck the error log for details."))
    
    def display_preview(self, image):
        """Display preview image on canvas"""
        # Resize to fit canvas
        canvas_width = self.preview_canvas.winfo_width()
        canvas_height = self.preview_canvas.winfo_height()
        
        if canvas_width < 100:  # Canvas not initialized yet
            canvas_width = 800
            canvas_height = 600
        
        # Calculate scaling
        img_ratio = image.width / image.height
        canvas_ratio = canvas_width / canvas_height
        
        if img_ratio > canvas_ratio:
            new_width = canvas_width - 40
            new_height = int(new_width / img_ratio)
        else:
            new_height = canvas_height - 40
            new_width = int(new_height * img_ratio)
        
        resized = image.resize((new_width, new_height), Image.Resampling.LANCZOS)
        photo = ImageTk.PhotoImage(resized)
        
        # Display on canvas
        self.preview_canvas.delete("all")
        self.preview_canvas.create_image(canvas_width // 2, canvas_height // 2, image=photo)
        self.preview_canvas.image = photo  # Keep reference
    
    def start_scan(self):
        """Start a single scan"""
        if not TWAIN_AVAILABLE or not self.scanner:
            messagebox.showinfo("Demo Mode", "Scanning would occur here when scanner is connected")
            return
        
        self.status_label.config(text="Scanning...", fg='#ffa500')
        self.scan_btn.config(state=tk.DISABLED)
        
        # WIA scanners need UI on main thread
        if self.is_wia or 'WIA-' in self.scanner_name:
            self._do_scan()
        else:
            threading.Thread(target=self._do_scan, daemon=True).start()
    
    def _do_scan(self):
        """Perform scan in background thread"""
        try:
            self.logger.info(f"Starting scan: {self.resolution.get()} DPI, {self.color_mode.get()} mode")
            
            if not self.scanner:
                raise Exception("Scanner not initialized")
            
            # Validate output directory
            if not os.path.exists(self.output_dir.get()):
                self.logger.info(f"Creating output directory: {self.output_dir.get()}")
                os.makedirs(self.output_dir.get(), exist_ok=True)
            
            # Configure scanner
            resolution = self.resolution.get()
            
            if resolution < 75 or resolution > 6400:
                raise ValueError(f"Invalid resolution: {resolution}. Must be between 75 and 6400 DPI.")
            
            self.logger.info(f"Scanner: {self.scanner_name}, is_wia: {self.is_wia}")
            
            # WIA drivers have limited capability support
            if self.is_wia or 'WIA-' in self.scanner_name:
                self.logger.warning("WIA scanner detected - using simplified settings")
                self.logger.info("Note: Resolution, color mode, and film settings will be set through scanner UI")
                self.logger.info("IMPORTANT: In the Epson dialog, select 'Film' or 'Transparency' as document type!")
            else:
                try:
                    # Configure for film/transparency scanning
                    try:
                        # Disable feeder (use flatbed/transparency unit)
                        self.scanner.SetCapability(twain.CAP_FEEDERENABLED, twain.TWTY_BOOL, False)
                    except:
                        pass
                    
                    try:
                        # Set to transparency mode (enables backlight)
                        # ICAP_LIGHTPATH = 0x1005, TWLP_TRANSMISSIVE = 1
                        self.scanner.SetCapability(0x1005, twain.TWTY_UINT16, 1)
                        self.logger.info("Transparency/film mode enabled (backlight on)")
                    except:
                        self.logger.warning("Could not set transparency mode - set 'Film' in scanner UI")
                    
                    # Set resolution
                    self.scanner.SetCapability(twain.ICAP_XRESOLUTION, twain.TWTY_FIX32, resolution)
                    self.scanner.SetCapability(twain.ICAP_YRESOLUTION, twain.TWTY_FIX32, resolution)
                    
                    # Set color mode
                    if self.color_mode.get() == "Color":
                        pixel_type = twain.TWPT_RGB
                    elif self.color_mode.get() == "Grayscale":
                        pixel_type = twain.TWPT_GRAY
                    else:
                        pixel_type = twain.TWPT_BW
                    
                    self.scanner.SetCapability(twain.ICAP_PIXELTYPE, twain.TWTY_UINT16, pixel_type)
                except Exception as e:
                    self.logger.warning(f"Could not set capabilities: {str(e)}. Using scanner defaults.")
            
            # Acquire image - Always show UI for film scanning
            self.logger.debug("Requesting image acquisition...")
            self.logger.info("Showing scanner UI for film settings")
            
            # Show instruction for film scanning
            messagebox.showinfo("Film Scanning - IMPORTANT!",
                "The Epson Scan window will open.\n\n"
                "CRITICAL STEPS:\n"
                "1. Look for 'Document Type' or mode selector\n"
                "2. Change from 'Reflective' to 'Film' or 'Transparency'\n"
                "3. Select 'Negative Film' or 'Positive Film'\n"
                "4. Choose film holder type (35mm, etc.)\n"
                "5. Set your scan resolution and area\n\n"
                "The backlight will turn ON when you select Film mode!")
            
            self.scanner.RequestAcquire(1, 1)  # Always show UI for film scanning
            
            # Get image data
            rv = self.scanner.XferImageNatively()
            
            # Handle both return formats: (handle, more) or just handle
            if isinstance(rv, tuple):
                image_handle = rv[0]
            else:
                image_handle = rv
            
            if not image_handle:
                raise Exception("No image data received from scanner")
            
            # Convert DIB handle to PIL Image using pytwain's method
            import tempfile
            temp_bmp = tempfile.mktemp(suffix='.bmp')
            
            try:
                # Use twain module's DIBToBMFile if available
                import twain
                if hasattr(twain, 'DIBToBMFile'):
                    twain.DIBToBMFile(image_handle, temp_bmp)
                    image = Image.open(temp_bmp)
                else:
                    # Fallback: save handle as temp file and open
                    # This is what pytwain does internally
                    from PIL import ImageWin
                    image = ImageWin.Dib(image_handle).image
            finally:
                if os.path.exists(temp_bmp):
                    try:
                        os.remove(temp_bmp)
                    except:
                        pass
            
            if image.size[0] == 0 or image.size[1] == 0:
                raise Exception("Invalid image dimensions received")
            
            self.logger.info(f"Image acquired: {image.size}, mode: {image.mode}")
            
            # Apply adjustments to scanned image
            image = self.apply_all_transforms(image)
            
            # Auto-detect film frames if enabled
            if self.auto_detect.get():
                self.logger.debug("Attempting frame detection...")
                frames = self.detect_film_frames(image)
                if frames:
                    self.logger.info(f"Detected {len(frames)} frames")
                    self.root.after(0, lambda: self.status_label.config(
                        text=f"Detected {len(frames)} frames, saving...", fg='#ffa500'))
                    self.save_detected_frames(image, frames)
                    return
                else:
                    self.logger.debug("No frames detected, saving as single image")
            
            # Save single image (no auto-detect or no frames found)
            filename = self.generate_filename()
            filepath = os.path.join(self.output_dir.get(), filename)
            
            os.makedirs(self.output_dir.get(), exist_ok=True)
            
            # Validate filepath
            if len(filepath) > 260:  # Windows path length limit
                raise Exception("File path too long. Choose a shorter output directory or filename.")
            
            self.logger.info(f"Saving image to: {filepath}")
            
            if self.file_format.get() == "TIFF":
                image.save(filepath, "TIFF", compression="tiff_lzw")
            elif self.file_format.get() == "PNG":
                image.save(filepath, "PNG")
            else:
                image.save(filepath, "JPEG", quality=95)
            
            # Verify file was created
            if not os.path.exists(filepath):
                raise Exception("File was not created successfully")
            
            file_size = os.path.getsize(filepath) / (1024 * 1024)  # MB
            self.logger.info(f"Scan completed successfully: {file_size:.2f} MB")
            
            self.scanned_images.append(filepath)
            self.root.after(0, self.scan_complete, filepath)
            
        except ValueError as e:
            error_msg = f"Invalid setting: {str(e)}"
            self.logger.error(error_msg)
            self.root.after(0, lambda: self.scan_failed(error_msg))
            
        except PermissionError as e:
            error_msg = f"Permission denied: Cannot write to output directory"
            self.logger.error(f"{error_msg}: {str(e)}")
            self.root.after(0, lambda: self.scan_failed(error_msg))
            
        except Exception as e:
            error_msg = f"Scan error: {str(e)}"
            self.logger.error(f"{error_msg}\n{traceback.format_exc()}")
            self.root.after(0, lambda: self.scan_failed(error_msg))
    
    def apply_all_transforms(self, image):
        """Apply all transformations and adjustments to an image"""
        img = image.copy()
        
        # Apply rotation
        if self.rotation_angle.get() != 0:
            img = img.rotate(-self.rotation_angle.get(), expand=True)
        
        # Apply flips
        if self.flip_horizontal.get():
            img = ImageOps.mirror(img)
        
        if self.flip_vertical.get():
            img = ImageOps.flip(img)
        
        # Apply adjustments
        img = self.apply_adjustments(img)
        
        return img
    
    def generate_filename(self):
        """Generate filename for scanned image"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        ext_map = {"TIFF": "tif", "PNG": "png", "JPEG": "jpg"}
        ext = ext_map[self.file_format.get()]
        
        if self.auto_increment.get():
            filename = f"film_scan_{self.scan_counter:04d}.{ext}"
            self.scan_counter += 1
        else:
            filename = f"film_scan_{timestamp}.{ext}"
        
        return filename
    
    def detect_film_frames(self, image):
        """Detect individual film frames in a scanned image"""
        try:
            # Convert to grayscale for analysis
            gray = ImageOps.grayscale(image)
            
            # Convert to numpy array
            img_array = np.array(gray)
            
            # Find bright areas (film frames are typically brighter than background)
            # Threshold to create binary image
            threshold = np.mean(img_array) + np.std(img_array) * 0.5
            binary = img_array > threshold
            
            # Find contiguous regions (simple row/column analysis)
            frames = []
            min_frame_size = 1000  # Minimum pixels for a frame
            
            # Detect columns with content
            col_sums = np.sum(binary, axis=0)
            col_threshold = img_array.shape[0] * 0.1  # At least 10% of column height
            
            in_frame = False
            frame_start = 0
            
            for i, col_sum in enumerate(col_sums):
                if col_sum > col_threshold and not in_frame:
                    frame_start = i
                    in_frame = True
                elif col_sum <= col_threshold and in_frame:
                    # Found end of frame
                    frame_width = i - frame_start
                    if frame_width > 100:  # Minimum width
                        # Now find top and bottom of frame
                        frame_region = binary[:, frame_start:i]
                        row_sums = np.sum(frame_region, axis=1)
                        row_threshold = frame_width * 0.1
                        
                        rows_with_content = np.where(row_sums > row_threshold)[0]
                        if len(rows_with_content) > 100:  # Minimum height
                            top = rows_with_content[0]
                            bottom = rows_with_content[-1]
                            
                            # Add some padding
                            padding = 10
                            left = max(0, frame_start - padding)
                            right = min(image.width, i + padding)
                            top = max(0, top - padding)
                            bottom = min(image.height, bottom + padding)
                            
                            frames.append((left, top, right, bottom))
                    
                    in_frame = False
            
            return frames
            
        except Exception as e:
            print(f"Error detecting frames: {e}")
            return []
    
    def save_detected_frames(self, source_image, frames):
        """Save individual detected frames"""
        try:
            saved_files = []
            
            for i, (left, top, right, bottom) in enumerate(frames):
                # Crop the frame
                frame_img = source_image.crop((left, top, right, bottom))
                
                if frame_img:
                    filename = self.generate_filename()
                    filepath = os.path.join(self.output_dir.get(), filename)
                    
                    os.makedirs(self.output_dir.get(), exist_ok=True)
                    
                    if self.file_format.get() == "TIFF":
                        frame_img.save(filepath, "TIFF", compression="tiff_lzw")
                    elif self.file_format.get() == "PNG":
                        frame_img.save(filepath, "PNG")
                    else:
                        frame_img.save(filepath, "JPEG", quality=95)
                    
                    saved_files.append(filepath)
                    self.scanned_images.append(filepath)
            
            if saved_files:
                self.root.after(0, lambda: self.multi_scan_complete(saved_files))
            else:
                self.root.after(0, lambda: self.scan_failed("No frames detected"))
                
        except Exception as e:
            self.root.after(0, lambda: self.scan_failed(f"Error saving frames: {str(e)}"))
    
    def multi_scan_complete(self, filepaths):
        """Handle successful multi-frame scan"""
        self.status_label.config(text=f"Saved {len(filepaths)} frames!", fg='#00ff00')
        self.scan_btn.config(state=tk.NORMAL)
        self.stats_label.config(text=f"Scans completed: {len(self.scanned_images)}")
        files_list = "\n".join([os.path.basename(f) for f in filepaths[:10]])
        if len(filepaths) > 10:
            files_list += f"\n... and {len(filepaths) - 10} more"
        messagebox.showinfo("Scan Complete", f"Saved {len(filepaths)} frames:\n{files_list}")
    
    def scan_complete(self, filepath):
        """Handle successful scan"""
        self.status_label.config(text="Scan complete!", fg='#00ff00')
        self.scan_btn.config(state=tk.NORMAL)
        self.stats_label.config(text=f"Scans completed: {len(self.scanned_images)}")
        messagebox.showinfo("Scan Complete", f"Image saved to:\n{filepath}")
    
    def scan_failed(self, error):
        """Handle scan failure"""
        self.status_label.config(text=f"Scan failed", fg='#ff4444')
        self.scan_btn.config(state=tk.NORMAL)
        
        error_details = f"Error during scanning:\n{error}\n\n"
        error_details += "Troubleshooting tips:\n"
        error_details += "• Check scanner is connected and powered on\n"
        error_details += "• Ensure scanner driver is installed\n"
        error_details += "• Try lowering the resolution\n"
        error_details += "• Check output directory permissions\n"
        error_details += "• View error log for more details\n"
        
        messagebox.showerror("Scan Failed", error_details)
        self.logger.error(f"Scan failed: {error}")
    
    def batch_scan(self):
        """Start batch scanning mode"""
        if not TWAIN_AVAILABLE or not self.scanner:
            messagebox.showinfo("Demo Mode", "Batch scanning would occur here when scanner is connected")
            return
        
        count = simpledialog.askinteger("Batch Scan", "How many images to scan?", 
                                          initialvalue=5, minvalue=1, maxvalue=100)
        if count:
            self.status_label.config(text=f"Batch scanning {count} images...", fg='#ffa500')
            threading.Thread(target=self._do_batch_scan, args=(count,), daemon=True).start()
    
    def _do_batch_scan(self, count):
        """Perform batch scan"""
        for i in range(count):
            self.root.after(0, lambda i=i: self.status_label.config(
                text=f"Scanning {i+1} of {count}...", fg='#ffa500'))
            self._do_scan()
            if i < count - 1:
                self.root.after(0, lambda: messagebox.showinfo("Ready", "Load next film and click OK"))
        
        self.root.after(0, lambda: self.status_label.config(text="Batch scan complete!", fg='#00ff00'))
    
    def add_to_queue(self):
        """Add current settings to scan queue"""
        try:
            # Validate settings before adding
            if self.resolution.get() < 75 or self.resolution.get() > 6400:
                messagebox.showerror("Invalid Settings", "Resolution must be between 75 and 6400 DPI")
                return
            
            if not os.path.exists(self.output_dir.get()):
                if not messagebox.askyesno("Create Directory", 
                    f"Output directory does not exist:\n{self.output_dir.get()}\n\nCreate it now?"):
                    return
                os.makedirs(self.output_dir.get(), exist_ok=True)
            
            queue_item = {
                'resolution': self.resolution.get(),
                'color_mode': self.color_mode.get(),
                'file_format': self.file_format.get(),
                'brightness': self.brightness.get(),
                'contrast': self.contrast.get(),
                'exposure': self.exposure.get(),
                'invert_negative': self.invert_negative.get(),
                'remove_dust': self.remove_dust.get(),
                'rotation_angle': self.rotation_angle.get(),
                'flip_horizontal': self.flip_horizontal.get(),
                'flip_vertical': self.flip_vertical.get(),
                'auto_detect': self.auto_detect.get(),
                'timestamp': datetime.now().strftime("%H:%M:%S")
            }
            
            self.scan_queue.append(queue_item)
            self.logger.info(f"Added scan to queue: {len(self.scan_queue)} items total")
            self.update_queue_display()
            messagebox.showinfo("Added to Queue", 
                              f"Scan added to queue with current settings\n"
                              f"Resolution: {queue_item['resolution']} DPI\n"
                              f"Queue position: {len(self.scan_queue)}")
            
        except Exception as e:
            error_msg = f"Could not add to queue: {str(e)}"
            self.logger.error(error_msg)
            messagebox.showerror("Queue Error", error_msg)
    
    def clear_queue(self):
        """Clear all items from queue"""
        if not self.scan_queue:
            messagebox.showinfo("Queue Empty", "No items in queue to clear")
            return
        
        if messagebox.askyesno("Clear Queue", f"Clear all {len(self.scan_queue)} items from queue?"):
            self.scan_queue.clear()
            self.update_queue_display()
    
    def update_queue_display(self):
        """Update queue status display"""
        count = len(self.scan_queue)
        self.queue_label.config(text=f"Queue: {count} scan{'s' if count != 1 else ''}")
        
        if count > 0 and not self.queue_processing:
            self.process_queue_btn.config(state=tk.NORMAL)
        else:
            self.process_queue_btn.config(state=tk.DISABLED if self.queue_processing else tk.NORMAL)
    
    def process_queue(self):
        """Start processing the scan queue"""
        if not self.scan_queue:
            messagebox.showinfo("Queue Empty", "No scans in queue to process")
            return
        
        if not TWAIN_AVAILABLE or not self.scanner:
            messagebox.showinfo("Demo Mode", "Queue processing would occur here when scanner is connected")
            return
        
        self.queue_processing = True
        self.queue_paused = False
        self.process_queue_btn.config(state=tk.DISABLED)
        self.pause_queue_btn.config(state=tk.NORMAL)
        self.scan_btn.config(state=tk.DISABLED)
        self.batch_btn.config(state=tk.DISABLED)
        
        self.status_label.config(text=f"Processing queue: {len(self.scan_queue)} items...", fg='#ffa500')
        threading.Thread(target=self._process_queue_thread, daemon=True).start()
    
    def toggle_pause_queue(self):
        """Pause or resume queue processing"""
        self.queue_paused = not self.queue_paused
        
        if self.queue_paused:
            self.pause_queue_btn.config(text="▶ Resume Queue", bg='#28a745')
            self.status_label.config(text="Queue paused", fg='#ffa500')
        else:
            self.pause_queue_btn.config(text="⏸ Pause Queue", bg='#ffc107')
            self.status_label.config(text="Queue processing...", fg='#ffa500')
    
    def _process_queue_thread(self):
        """Process all items in the queue"""
        total = len(self.scan_queue)
        completed = 0
        
        while self.scan_queue and self.queue_processing:
            # Wait if paused
            while self.queue_paused:
                threading.Event().wait(0.5)
                if not self.queue_processing:  # Check if stopped during pause
                    break
            
            if not self.queue_processing:
                break
            
            # Get next item from queue
            queue_item = self.scan_queue[0]
            completed += 1
            
            self.root.after(0, lambda c=completed, t=total: self.status_label.config(
                text=f"Queue: Scanning {c} of {t}...", fg='#ffa500'))
            
            # Apply settings from queue item
            self.root.after(0, lambda item=queue_item: self._apply_queue_settings(item))
            threading.Event().wait(0.5)  # Give UI time to update
            
            # Perform the scan
            try:
                self._do_scan_from_queue(queue_item)
                
                # Remove completed item
                self.scan_queue.pop(0)
                self.root.after(0, self.update_queue_display)
                
                # Prompt to load next film if more items in queue
                if self.scan_queue:
                    self.root.after(0, lambda c=completed, t=total: 
                                  messagebox.showinfo("Next Scan", 
                                                    f"Completed {c} of {t}\n\n"
                                                    "Load next film and click OK to continue"))
            
            except Exception as e:
                self.root.after(0, lambda err=str(e): 
                              messagebox.showerror("Queue Error", 
                                                 f"Error during queue processing:\n{err}\n\nQueue stopped."))
                break
        
        # Queue finished
        self.queue_processing = False
        self.root.after(0, self._queue_complete, completed)
    
    def _apply_queue_settings(self, queue_item):
        """Apply settings from queue item to current settings"""
        self.resolution.set(queue_item['resolution'])
        self.color_mode.set(queue_item['color_mode'])
        self.file_format.set(queue_item['file_format'])
        self.brightness.set(queue_item['brightness'])
        self.contrast.set(queue_item['contrast'])
        self.exposure.set(queue_item['exposure'])
        self.invert_negative.set(queue_item['invert_negative'])
        self.remove_dust.set(queue_item['remove_dust'])
        self.rotation_angle.set(queue_item['rotation_angle'])
        self.flip_horizontal.set(queue_item['flip_horizontal'])
        self.flip_vertical.set(queue_item['flip_vertical'])
        self.auto_detect.set(queue_item['auto_detect'])
    
    def _do_scan_from_queue(self, queue_item):
        """Perform scan with queue item settings (synchronous)"""
        try:
            # Configure scanner
            resolution = queue_item['resolution']
            self.scanner.SetCapability(twain.ICAP_XRESOLUTION, twain.TWTY_FIX32, resolution)
            self.scanner.SetCapability(twain.ICAP_YRESOLUTION, twain.TWTY_FIX32, resolution)
            
            # Set color mode
            if queue_item['color_mode'] == "Color":
                pixel_type = twain.TWPT_RGB
            elif queue_item['color_mode'] == "Grayscale":
                pixel_type = twain.TWPT_GRAY
            else:
                pixel_type = twain.TWPT_BW
            
            self.scanner.SetCapability(twain.ICAP_PIXELTYPE, twain.TWTY_UINT16, pixel_type)
            
            # Acquire image
            self.scanner.RequestAcquire(0, 0)
            image_data = self.scanner.XferImageNatively()[0]
            image = Image.open(image_data)
            
            # Apply transformations (using queue settings)
            image = self._apply_transforms_from_queue(image, queue_item)
            
            # Auto-detect or save
            if queue_item['auto_detect']:
                frames = self.detect_film_frames(image)
                if frames:
                    self._save_frames_sync(image, frames)
                    return
            
            # Save single image
            filename = self.generate_filename()
            filepath = os.path.join(self.output_dir.get(), filename)
            os.makedirs(self.output_dir.get(), exist_ok=True)
            
            if queue_item['file_format'] == "TIFF":
                image.save(filepath, "TIFF", compression="tiff_lzw")
            elif queue_item['file_format'] == "PNG":
                image.save(filepath, "PNG")
            else:
                image.save(filepath, "JPEG", quality=95)
            
            self.scanned_images.append(filepath)
            
        except Exception as e:
            raise Exception(f"Scan failed: {str(e)}")
    
    def _apply_transforms_from_queue(self, image, queue_item):
        """Apply transformations from queue settings"""
        img = image.copy()
        
        # Rotation
        if queue_item['rotation_angle'] != 0:
            img = img.rotate(-queue_item['rotation_angle'], expand=True)
        
        # Flips
        if queue_item['flip_horizontal']:
            img = ImageOps.mirror(img)
        if queue_item['flip_vertical']:
            img = ImageOps.flip(img)
        
        # Adjustments
        if queue_item['brightness'] != 1.0:
            enhancer = ImageEnhance.Brightness(img)
            img = enhancer.enhance(queue_item['brightness'])
        
        if queue_item['contrast'] != 1.0:
            enhancer = ImageEnhance.Contrast(img)
            img = enhancer.enhance(queue_item['contrast'])
        
        if queue_item['exposure'] != 0.0:
            exposure_factor = 1.0 + queue_item['exposure']
            img_array = np.array(img).astype(np.float32)
            img_array = np.clip(img_array * exposure_factor, 0, 255)
            img = Image.fromarray(img_array.astype(np.uint8))
        
        if queue_item['invert_negative']:
            if img.mode in ['RGB', 'L']:
                img = ImageOps.invert(img)
        
        if queue_item['remove_dust']:
            img = img.filter(ImageFilter.MedianFilter(size=3))
        
        return img
    
    def _save_frames_sync(self, source_image, frames):
        """Save frames synchronously (for queue processing)"""
        for left, top, right, bottom in frames:
            frame_img = source_image.crop((left, top, right, bottom))
            filename = self.generate_filename()
            filepath = os.path.join(self.output_dir.get(), filename)
            os.makedirs(self.output_dir.get(), exist_ok=True)
            
            if self.file_format.get() == "TIFF":
                frame_img.save(filepath, "TIFF", compression="tiff_lzw")
            elif self.file_format.get() == "PNG":
                frame_img.save(filepath, "PNG")
            else:
                frame_img.save(filepath, "JPEG", quality=95)
            
            self.scanned_images.append(filepath)
    
    def _queue_complete(self, completed):
        """Handle queue completion"""
        self.queue_processing = False
        self.process_queue_btn.config(state=tk.NORMAL if self.scan_queue else tk.DISABLED)
        self.pause_queue_btn.config(state=tk.DISABLED, text="⏸ Pause Queue", bg='#ffc107')
        self.scan_btn.config(state=tk.NORMAL)
        self.batch_btn.config(state=tk.NORMAL)
        self.status_label.config(text=f"Queue complete! {completed} scans processed", fg='#00ff00')
        self.stats_label.config(text=f"Scans completed: {len(self.scanned_images)}")
        self.update_queue_display()
        
        messagebox.showinfo("Queue Complete", 
                          f"All queue items processed!\n\n"
                          f"Completed: {completed} scans\n"
                          f"Total session scans: {len(self.scanned_images)}")


def main():
    root = tk.Tk()
    app = FilmScannerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()

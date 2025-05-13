import io
import os
import threading
from math import ceil
import fitz                # PyMuPDF
from PIL import Image, ImageOps
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.utils import ImageReader
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, BooleanVar
import webbrowser
import tempfile
import urllib.request
import shutil
import subprocess
import platform



class ThemeManager:
    """Handles theming for the application"""
    
    LIGHT_THEME = {
        "bg": "#f0f0f0",
        "fg": "#333333",
        "accent": "#1976d2",
        "button_bg": "#e0e0e0",
        "hover_bg": "#d0d0d0",
        "progress_bg": "#bbdefb",
        "progress_fg": "#1976d2",
        "frame_bg": "#ffffff",
        "status_good": "#4caf50",
        "status_warning": "#ff9800",
        "status_error": "#f44336",
    }
    
    DARK_THEME = {
        "bg": "#121212",
        "fg": "#e0e0e0",
        "accent": "#90caf9",
        "button_bg": "#333333",
        "hover_bg": "#424242",
        "progress_bg": "#424242",
        "progress_fg": "#90caf9",
        "frame_bg": "#1e1e1e",
        "status_good": "#81c784",
        "status_warning": "#ffb74d",
        "status_error": "#e57373",
    }
    
    @staticmethod
    def apply_theme(root, style, is_dark=True):
        theme = ThemeManager.DARK_THEME if is_dark else ThemeManager.LIGHT_THEME
        
        # Configure ttk style
        style.configure("TFrame", background=theme["bg"])
        style.configure("TLabel", background=theme["bg"], foreground=theme["fg"])
        style.configure("TButton", 
                        background=theme["button_bg"], 
                        foreground=theme["fg"],
                        focuscolor=theme["accent"])
        
        style.map("TButton",
                 background=[("active", theme["hover_bg"])],
                 foreground=[("active", theme["fg"])])
                 
        style.configure("TCheckbutton", 
                        background=theme["bg"], 
                        foreground=theme["fg"])
        
        style.configure("TLabelframe", 
                        background=theme["frame_bg"],
                        foreground=theme["fg"])
                        
        style.configure("TLabelframe.Label", 
                        background=theme["bg"],
                        foreground=theme["fg"],
                        font=("Roboto", 9, "bold"))
                        
        style.configure("TProgressbar", 
                        background=theme["progress_fg"],
                        troughcolor=theme["progress_bg"])
                        
        style.configure("TCombobox", 
                        background=theme["button_bg"],
                        fieldbackground=theme["frame_bg"],
                        foreground=theme["fg"])
                        
        # Configure root window
        root.configure(background=theme["bg"])
        
        # Configure standard tkinter widgets
        root.option_add("*Background", theme["bg"])
        root.option_add("*Foreground", theme["fg"])
        root.option_add("*selectBackground", theme["accent"])
        root.option_add("*selectForeground", theme["fg"])
        
        return theme

class AnimatedGif:
    """Class to handle animated GIF display"""
    
    def __init__(self, master, path, width=None, height=None, loop=True):
        self.master = master
        self.path = path
        self.loop = loop
        self.frames = []
        self.current_frame = 0
        self.playing = False
        self._load_frames(width, height)
        
        self.canvas = tk.Canvas(master, bd=0, highlightthickness=0)
        self.canvas_obj = None
        
    def _load_frames(self, width=None, height=None):
        try:
            gif = Image.open(self.path)
            self.frames = []
            self.delays = []
            
            try:
                while True:
                    # Get frame duration in milliseconds
                    delay = gif.info.get('duration', 100)  # Default to 100ms
                    
                    # Copy and resize the frame if dimensions provided
                    frame = gif.copy()
                    if width and height:
                        frame = frame.resize((width, height), Image.LANCZOS)
                    
                    # Convert to PhotoImage and store
                    photoframe = tk.PhotoImage(data=self._get_gif_frame_as_data(frame))
                    self.frames.append(photoframe)
                    self.delays.append(delay)
                    
                    # Move to next frame
                    gif.seek(gif.tell() + 1)
            except EOFError:
                pass  # End of frames
                
        except Exception as e:
            print(f"Error loading animation: {e}")
            self.frames = []
    
    def _get_gif_frame_as_data(self, img):
        """Convert PIL Image to a data string that PhotoImage can use"""
        with io.BytesIO() as buffer:
            img.save(buffer, format="gif")
            return buffer.getvalue()
    
    def pack(self, **kwargs):
        self.canvas.pack(**kwargs)
        self.canvas.config(width=self.frames[0].width(), height=self.frames[0].height())
        self.canvas_obj = self.canvas.create_image(0, 0, anchor=tk.NW, image=self.frames[0])
        
    def grid(self, **kwargs):
        self.canvas.grid(**kwargs)
        self.canvas.config(width=self.frames[0].width(), height=self.frames[0].height())
        self.canvas_obj = self.canvas.create_image(0, 0, anchor=tk.NW, image=self.frames[0])
    
    def place(self, **kwargs):
        self.canvas.place(**kwargs)
        self.canvas.config(width=self.frames[0].width(), height=self.frames[0].height())
        self.canvas_obj = self.canvas.create_image(0, 0, anchor=tk.NW, image=self.frames[0])
        
    def start(self):
        if not self.frames:
            return
            
        self.playing = True
        self._animate()
        
    def stop(self):
        self.playing = False
        
    def _animate(self):
        if not self.playing:
            return
            
        if self.frames:
            # Update to the next frame
            self.current_frame = (self.current_frame + 1) % len(self.frames)
            self.canvas.itemconfig(self.canvas_obj, image=self.frames[self.current_frame])
            
            # Schedule the next frame update
            delay = self.delays[self.current_frame]
            self.master.after(delay, self._animate)
            
            # If we've reached the end and not looping, stop
            if self.current_frame == len(self.frames) - 1 and not self.loop:
                self.playing = False

class SplashScreen(tk.Toplevel):
    """Custom splash screen with animation support"""
    
    def __init__(self, parent, animation_path="yy.gif", duration=3000):
        super().__init__(parent)
        self.parent = parent
        self.duration = duration
        
        # Configure window
        self.overrideredirect(True)  # No window decorations
        self.attributes('-topmost', True)
        
        # Set window size and position
        width, height = 720, 300
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - width) // 2
        y = (screen_height - height) // 2
        self.geometry(f"{width}x{height}+{x}+{y}")
        
        # Create a frame for the content
        self.frame = ttk.Frame(self)
        self.frame.pack(fill=tk.BOTH, expand=True)
        
        # Load the default animation if needed
        if not animation_path or not os.path.exists(animation_path):
            # Create a placeholder animation
            self.create_default_animation()
            
        # Add animation if available
        if animation_path and os.path.exists(animation_path):
            # Check if it's a webm, convert if needed
            if animation_path.lower().endswith('.webm'):
                gif_path = self.convert_webm_to_gif(animation_path)
                if gif_path:
                    animation_path = gif_path
                else:
                    self.create_default_animation()
                    animation_path = None
            
            if animation_path and animation_path.lower().endswith('.gif'):
                self.animation = AnimatedGif(self.frame, animation_path, width=200, height=200)
                self.animation.pack(pady=20)
                self.animation.start()
        
        # Application name
        ttk.Label(self.frame, text="Slide2Print: Convert Slides into Printable PDFs", 
                 font=("Roboto", 18, "bold")).pack(pady=10)
        
        # Progress bar for loading indication
        self.progress = ttk.Progressbar(self.frame, mode='indeterminate')
        self.progress.pack(fill=tk.X, padx=50, pady=10)
        self.progress.start(10)
        
        # Schedule closing
        self.after(duration, self.finish)
    
    def create_default_animation(self):
        """Create a placeholder animation label"""
        loading_frame = ttk.Frame(self.frame)
        loading_frame.pack(pady=20)
        
        # Create loading dots
        self.dots = []
        for i in range(5):
            dot = ttk.Label(loading_frame, text="●", font=("Arial", 24))
            dot.grid(row=0, column=i, padx=5)
            self.dots.append(dot)
        
        # Start dot animation
        self.animate_dots()
    
    def animate_dots(self, idx=0):
        """Animate the loading dots"""
        for i, dot in enumerate(self.dots):
            if i == idx:
                dot.configure(foreground="#90caf9")  # Highlight current dot
            else:
                dot.configure(foreground="#424242")  # Dim other dots
        
        next_idx = (idx + 1) % len(self.dots)
        self.after(200, lambda: self.animate_dots(next_idx))
    
    def convert_webm_to_gif(self, webm_path):
        """Convert webm to gif using ffmpeg if available"""
        try:
            # Check if ffmpeg is available
            try:
                subprocess.run(["ffmpeg", "-version"], 
                              stdout=subprocess.PIPE, 
                              stderr=subprocess.PIPE, 
                              check=True)
            except (subprocess.SubprocessError, FileNotFoundError):
                print("FFmpeg not available, can't convert webm to gif")
                return None
                
            # Create temporary gif path
            temp_dir = tempfile.gettempdir()
            gif_path = os.path.join(temp_dir, "splash_animation.gif")
            
            # Convert webm to gif using ffmpeg
            cmd = [
                "ffmpeg", "-y",
                "-i", webm_path,
                "-vf", "fps=10,scale=400:-1:flags=lanczos,split[s0][s1];[s0]palettegen[p];[s1][p]paletteuse",
                "-loop", "0",
                gif_path
            ]
            
            subprocess.run(cmd, 
                          stdout=subprocess.PIPE, 
                          stderr=subprocess.PIPE, 
                          check=True)
                          
            return gif_path if os.path.exists(gif_path) else None
            
        except Exception as e:
            print(f"Error converting webm to gif: {e}")
            return None

    def finish(self):
        """End the splash screen and show main window"""
        self.progress.stop()
        self.destroy()
        self.parent.deiconify()  # Show the main window

class MaterialButton(tk.Frame):
    """Custom material-style button with hover and ripple effects"""
    
    def __init__(self, master, text, command=None, width=None, height=None, 
                 bg="#1976d2", fg="white", hover_bg="#1565c0", **kwargs):
        super().__init__(master, **kwargs)
        
        self.bg = bg
        self.fg = fg
        self.hover_bg = hover_bg
        self.command = command
        
        # Set dimensions
        self.width = width or 100
        self.height = height or 36
        
        # Create canvas for button
        self.canvas = tk.Canvas(
            self, 
            width=self.width, 
            height=self.height,
            bg=self.bg,
            highlightthickness=0,
            bd=0
        )
        self.canvas.pack(fill=tk.BOTH, expand=True)
        
        # Draw button background
        self.bg_id = self.canvas.create_rectangle(
            0, 0, self.width, self.height,
            fill=self.bg, width=0
        )
        
        # Add text
        self.text_id = self.canvas.create_text(
            self.width//2, self.height//2,
            text=text,
            fill=self.fg,
            font=("Roboto", 10)
        )
        
        # Bind events
        self.canvas.bind("<Enter>", self._on_enter)
        self.canvas.bind("<Leave>", self._on_leave)
        self.canvas.bind("<Button-1>", self._on_press)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        
    def _on_enter(self, event):
        self.canvas.itemconfig(self.bg_id, fill=self.hover_bg)
    
    def _on_leave(self, event):
        self.canvas.itemconfig(self.bg_id, fill=self.bg)
    
    def _on_press(self, event):
        # Create ripple effect
        x, y = event.x, event.y
        self.ripple = self.canvas.create_oval(
            x-5, y-5, x+5, y+5,
            fill="#ffffff33",
            width=0
        )
        
        # Animate ripple
        self._animate_ripple(self.ripple, x, y)
    
    def _animate_ripple(self, ripple_id, x, y, size=5, alpha=51, step=1):
        if size < 350:  # Maximum ripple size
            # Expand ripple
            size += step * 5
            alpha = max(0, alpha - (step * 0.5))  # Fade out gradually
            
            # Update ripple
            self.canvas.coords(ripple_id, x-size, y-size, x+size, y+size)
            self.canvas.itemconfig(ripple_id, fill=f"#ffffff{alpha:02x}")
            
            # Continue animation
            self.after(10, lambda: self._animate_ripple(ripple_id, x, y, size, alpha, step))
        else:
            # Remove ripple when animation complete
            self.canvas.delete(ripple_id)
    
    def _on_release(self, event):
        if self.command:
            self.command()
    
    def config(self, **kwargs):
        """Update button configuration"""
        if "text" in kwargs:
            self.canvas.itemconfig(self.text_id, text=kwargs["text"])
        
        if "bg" in kwargs:
            self.bg = kwargs["bg"]
            self.canvas.itemconfig(self.bg_id, fill=self.bg)
            self.canvas.config(bg=self.bg)
            
        if "fg" in kwargs:
            self.fg = kwargs["fg"]
            self.canvas.itemconfig(self.text_id, fill=self.fg)
            
        if "hover_bg" in kwargs:
            self.hover_bg = kwargs["hover_bg"]
            
        if "state" in kwargs:
            if kwargs["state"] == "disabled":
                self.canvas.unbind("<Button-1>")
                self.canvas.unbind("<ButtonRelease-1>")
                self.canvas.itemconfig(self.bg_id, fill="#bbbbbb")
                self.canvas.itemconfig(self.text_id, fill="#888888")
            else:
                self.canvas.bind("<Button-1>", self._on_press)
                self.canvas.bind("<ButtonRelease-1>", self._on_release)
                self.canvas.itemconfig(self.bg_id, fill=self.bg)
                self.canvas.itemconfig(self.text_id, fill=self.fg)

class PDFProcessor:
    def __init__(self, input_path, output_path, skip_first=True, add_title=True, 
                 title_on_first_only=False, pages_per_sheet=3):
        self.input_path = input_path
        self.output_path = output_path
        self.skip_first = skip_first
        self.add_title = add_title
        self.title_on_first_only = title_on_first_only
        self.pages_per_sheet = pages_per_sheet

    def process(self, progress_callback=None):
        # Open PDF and get Title metadata or fallback to filename
        doc = fitz.open(self.input_path)
        raw_title = doc.metadata.get("title", "").strip()
        title = raw_title or os.path.splitext(os.path.basename(self.input_path))[0]

        # Calculate starting page and total pages to process
        start_page = 1 if self.skip_first else 0
        total_pages = doc.page_count - start_page
        
        if total_pages <= 0:
            doc.close()
            raise ValueError(f"'{os.path.basename(self.input_path)}' has no pages to process.")

        # Calculate how many output pages we'll need
        output_page_count = ceil(total_pages / self.pages_per_sheet)
        
        # Create a new PDF with reportlab
        c = canvas.Canvas(self.output_path, pagesize=A4)
        width_pt, height_pt = A4
        
        # Process all pages in groups
        for output_page in range(output_page_count):
            # Reset page for each new output page
            if output_page > 0:
                c.showPage()
                
            # Calculate which source pages go on this output page
            page_start_idx = start_page + (output_page * self.pages_per_sheet)
            page_end_idx = min(page_start_idx + self.pages_per_sheet, doc.page_count)
            current_page_count = page_end_idx - page_start_idx
            
            # Set up page layout
            margin = 20 * mm
            y_cursor = height_pt - margin
            
            # Add title if requested
            should_add_title = self.add_title and (output_page == 0 or not self.title_on_first_only)
            if should_add_title:
                c.setFont("Helvetica", 9)
                # Truncate title if too long
                max_title_width = width_pt - 2 * margin
                title_text = f"{title} - Sheet {output_page + 1}/{output_page_count}"
                
                # Measure text width and truncate if needed
                text_width = c.stringWidth(title_text, "Helvetica", 9)
                if text_width > max_title_width:
                    # Calculate how many characters we can fit
                    char_width = text_width / len(title_text)
                    max_chars = int(max_title_width / char_width) - 3  # -3 for ellipsis
                    truncated_title = title[:max_chars] + "..."
                    title_text = f"{truncated_title} - Sheet {output_page + 1}/{output_page_count}"
                
                c.drawString(20 * mm, height_pt - margin + 5 * mm, title_text)
            else:
                y_cursor = height_pt - 10 * mm  # Less margin if no title
            
            # Calculate section height based on number of images on this page
            section_h = (y_cursor - 10 * mm) / current_page_count
            
            # Process each page for this output sheet
            for i in range(current_page_count):
                src_idx = page_start_idx + i
                
                # Render and invert the page
                page = doc.load_page(src_idx)
                pix = page.get_pixmap(matrix=fitz.Matrix(2, 2), alpha=False)
                img = Image.frombytes("RGB", (pix.width, pix.height), pix.samples)
                img = ImageOps.invert(img)
                
                # Scale to fit width and section height
                scale = min((width_pt - 2*margin) / img.width, section_h / img.height)
                w, h = img.width * scale, img.height * scale
                x = (width_pt - w) / 2
                y = y_cursor - h
                
                # Draw the image
                buf = io.BytesIO()
                img.save(buf, format="PNG")
                buf.seek(0)
                ir = ImageReader(buf)
                c.drawImage(ir, x, y, width=w, height=h)
                
                # Add page number
                c.setFont("Helvetica", 8)
                c.drawString(width_pt - margin - 20, y, f"Page {src_idx + 1}")
                
                # Move cursor down for next image
                y_cursor = y - 5 * mm
                
                # Update progress
                if progress_callback:
                    progress_callback(src_idx - start_page + 1, total_pages)
        
        # Finish and save PDF
        c.save()
        doc.close()
        return output_page_count

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Slide2Print: Convert Slides into Printable PDFs")
        self.geometry("600x750")
        
        # Initialize variables
        self.file_paths = []
        self.output_dir = ""
        self.failures = []
        self.skip_first_var = BooleanVar(value=True)
        self.add_title_var = BooleanVar(value=True)
        self.title_on_first_only_var = BooleanVar(value=False)  # NEW option
        self.dark_mode_var = BooleanVar(value=True)  # Default to dark mode
        self.pages_per_sheet_var = tk.IntVar(value=3)
        self.animation_path = ""
        
        # Apply dark mode on startup
        self.style = ttk.Style()
        self.theme = ThemeManager.apply_theme(self, self.style, is_dark=True)
        
        # Hide main window initially
        self.withdraw()
        
        # Show splash screen
        self.splash = SplashScreen(self)
        
        # Create main frame when splash ends
        self.after(2000, self.setup_ui)
        
    def setup_ui(self):
        # Configure padding
        self.configure(padx=20, pady=20)
        
        # Create main frame
        main_frame = ttk.Frame(self)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Create menu bar
        self.create_menu()
        
        # Header with application name and theme toggle
        header_frame = ttk.Frame(main_frame)
        header_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(header_frame, text="Slide2Print: Convert Slides into Printable PDFs", 
                 font=("Roboto", 16, "bold")).pack(side=tk.LEFT)
        
        # Theme toggle
        theme_frame = ttk.Frame(header_frame)
        theme_frame.pack(side=tk.RIGHT)
        
        ttk.Label(theme_frame, text="Dark Mode").pack(side=tk.LEFT, padx=5)
        theme_toggle = ttk.Checkbutton(theme_frame, variable=self.dark_mode_var,
                                      command=self.toggle_theme)
        theme_toggle.pack(side=tk.LEFT)
        
        # File selection section
        file_frame = ttk.LabelFrame(main_frame, text="Input Files")
        file_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        file_btn_frame = ttk.Frame(file_frame)
        file_btn_frame.pack(fill=tk.X, pady=5)
        
        # Custom material buttons
        self.select_btn = MaterialButton(
            file_btn_frame, 
            text="Select PDF Files", 
            command=self.select_files,
            bg=self.theme["accent"],
            fg="white",
            hover_bg="#1565c0",
            width=150
        )
        self.select_btn.pack(side=tk.LEFT, padx=5)
        
        self.clear_btn = MaterialButton(
            file_btn_frame, 
            text="Clear Selection", 
            command=self.clear_files,
            bg="#e57373",
            fg="white",
            hover_bg="#ef5350",
            width=150
        )
        self.clear_btn.pack(side=tk.LEFT, padx=5)
        
        # Files list with scrollbar
        files_frame = ttk.Frame(file_frame)
        files_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=5)
        
        scrollbar = ttk.Scrollbar(files_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.files_list = tk.Listbox(files_frame, width=80, height=8,
                                     bg=self.theme["frame_bg"],
                                     fg=self.theme["fg"],
                                     selectbackground=self.theme["accent"])
        self.files_list.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.files_list.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=self.files_list.yview)
        
        # Output directory selection
        output_frame = ttk.LabelFrame(main_frame, text="Output Settings")
        output_frame.pack(fill=tk.X, pady=(0, 10))
        
        output_btn = MaterialButton(
            output_frame,
            text="Select Output Directory",
            command=self.select_directory,
            bg=self.theme["accent"],
            fg="white",
            hover_bg="#1565c0",
            width=180
        )
        output_btn.pack(anchor=tk.W, padx=5, pady=5)
        
        self.output_label = ttk.Label(output_frame, text="No output folder selected")
        self.output_label.pack(anchor=tk.W, padx=5, pady=5)
        
        # Options frame
        options_frame = ttk.LabelFrame(main_frame, text="Processing Options")
        options_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Left options column
        left_opts = ttk.Frame(options_frame)
        left_opts.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=5)
        
        # Checkboxes for options
        ttk.Checkbutton(left_opts, text="Skip first page", 
                       variable=self.skip_first_var).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(left_opts, text="Add title to sheets", 
                       variable=self.add_title_var).pack(anchor=tk.W, pady=2)
        ttk.Checkbutton(left_opts, text="Title on first page only", 
                       variable=self.title_on_first_only_var).pack(anchor=tk.W, pady=2)
        
        # Right options column
        right_opts = ttk.Frame(options_frame)
        right_opts.pack(side=tk.LEFT, fill=tk.Y, padx=10, pady=5)
        
        # Pages per sheet option
        sheet_frame = ttk.Frame(right_opts)
        sheet_frame.pack(anchor=tk.W, pady=2)
        
        ttk.Label(sheet_frame, text="Pages per sheet:").pack(side=tk.LEFT)
        pages_combobox = ttk.Combobox(sheet_frame, textvariable=self.pages_per_sheet_var, width=5)
        pages_combobox['values'] = (1, 2, 3, 4, 6)
        pages_combobox.pack(side=tk.LEFT, padx=5)
        pages_combobox.state(['readonly'])
        
        # Splash animation setting
        animation_frame = ttk.Frame(right_opts)
        animation_frame.pack(anchor=tk.W, pady=2, fill=tk.X)
        
        
        
        # Progress section
        progress_frame = ttk.LabelFrame(main_frame, text="Progress")
        progress_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Overall progress with label
        overall_frame = ttk.Frame(progress_frame)
        overall_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(overall_frame, text="Overall:").pack(side=tk.LEFT, padx=(0, 10))
        self.progress = ttk.Progressbar(overall_frame, orient='horizontal', mode='determinate')
        self.progress.pack(side=tk.LEFT, fill='x', expand=True)
        
        self.status_label = ttk.Label(progress_frame, text="Ready")
        self.status_label.pack(anchor=tk.W, padx=5, pady=2)
        
        # Current file progress with label
        detail_frame = ttk.Frame(progress_frame)
        detail_frame.pack(fill=tk.X, padx=5, pady=5)
        
        ttk.Label(detail_frame, text="Current File:").pack(side=tk.LEFT, padx=(0, 10))
        self.detail_progress = ttk.Progressbar(detail_frame, orient='horizontal', mode='determinate')
        self.detail_progress.pack(side=tk.LEFT, fill='x', expand=True)
        
        self.detail_label = ttk.Label(progress_frame, text="")
        self.detail_label.pack(anchor=tk.W, padx=5, pady=2)
        
        # Loading animation frame
        self.animation_frame = ttk.Frame(progress_frame)
        self.animation_frame.pack(fill=tk.X, padx=5, pady=5)
        self.loading_animation = None
        
        # Button frame - ensuring it's always visible
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=10, side=tk.BOTTOM)
        
        # Process button uses MaterialButton for better visibility
        self.process_btn = MaterialButton(
            button_frame, 
            text="Start Processing", 
            command=self.start_processing,
            bg="#4caf50",
            fg="white",
            hover_bg="#388e3c",
            width=180,
            height=40
        )
        self.process_btn.pack(side=tk.LEFT, padx=5)
        
        # Exit button
        exit_btn = MaterialButton(
            button_frame, 
            text="Exit", 
            command=self.quit,
            bg="#f44336",
            fg="white",
            hover_bg="#d32f2f",
            width=100,
            height=40
        )
        exit_btn.pack(side=tk.RIGHT, padx=5)

    def create_menu(self):
        """Create application menu"""
        menubar = tk.Menu(self)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="Select PDF Files", command=self.select_files)
        file_menu.add_command(label="Select Output Directory", command=self.select_directory)
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.quit)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Options menu
        options_menu = tk.Menu(menubar, tearoff=0)
        options_menu.add_checkbutton(label="Skip First Page", variable=self.skip_first_var)
        options_menu.add_checkbutton(label="Add Title", variable=self.add_title_var)
        options_menu.add_checkbutton(label="Title on First Page Only", variable=self.title_on_first_only_var)
        options_menu.add_separator()
        options_menu.add_checkbutton(label="Dark Mode", variable=self.dark_mode_var, 
                                    command=self.toggle_theme)
        menubar.add_cascade(label="Options", menu=options_menu)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        help_menu.add_command(label="About", command=self.show_about)
        help_menu.add_command(label="Help", command=self.show_help)
        menubar.add_cascade(label="Help", menu=help_menu)
        
        self.config(menu=menubar)

    def toggle_theme(self):
        """Toggle between light and dark theme"""
        is_dark = self.dark_mode_var.get()
        self.theme = ThemeManager.apply_theme(self, self.style, is_dark=is_dark)
        
        # Update specific widgets that need manual updating
        self.files_list.config(
            bg=self.theme["frame_bg"],
            fg=self.theme["fg"],
            selectbackground=self.theme["accent"]
        )
        
        # Update material buttons
        button_bg = self.theme["accent"]
        button_fg = "white" if is_dark else "white"
        hover_bg = "#1565c0" if is_dark else "#0d47a1"
        
        self.select_btn.config(bg=button_bg, fg=button_fg, hover_bg=hover_bg)
        self.clear_btn.config(bg="#e57373", fg="white", hover_bg="#ef5350")
        self.process_btn.config(bg="#4caf50", fg="white", hover_bg="#388e3c")

    def show_about(self):
        """Show about dialog"""
        messagebox.showinfo(
            "About Slide2Print: Convert Slides into Printable PDFs", 
            "Slide2Print: Convert Slides into Printable PDFs v2.0\n\n"
            "A modern tool for processing PDFs with a material design UI.\n\n"
            "Features:\n"
            "- Process multiple PDFs in batch\n"
            "- Skip first pages\n"
            "- Add titles with customization\n"
            "- Dark mode support\n"
            "- Custom animations"
        )
    
    def show_help(self):
        """Show help information"""
        messagebox.showinfo(
            "Help", 
            "How to use the Slide2Print: Convert Slides into Printable PDFs:\n\n"
            "1. Select PDF files using the 'Select PDF Files' button\n"
            "2. Choose an output directory\n"
            "3. Configure processing options\n"
            "4. Click 'Start Processing'\n\n"
            "Options:\n"
            "- Skip first page: Ignore the first page of each PDF\n"
            "- Add title: Add PDF title to each sheet\n"
            "- Title on first page only: Only add title to the first sheet\n"
            "- Pages per sheet: Number of pages to include on each output sheet"
        )

    def select_animation(self):
        """Select custom startup animation"""
        filetypes = [
            ("Animation Files", "*.gif;*.webm"),
            ("GIF Files", "*.gif"),
            ("WebM Files", "*.webm"),
            ("All Files", "*.*")
        ]
        path = filedialog.askopenfilename(
            title="Select Startup Animation",
            filetypes=filetypes
        )
        
        if path:
            self.animation_path = path
            filename = os.path.basename(path)
            messagebox.showinfo(
                "Animation Selected", 
                f"Animation '{filename}' selected.\nIt will be used the next time you start the application."
            )
            
            # Save animation path for next startup
            try:
                config_dir = os.path.join(os.path.expanduser("~"), ".pdf_processor")
                os.makedirs(config_dir, exist_ok=True)
                
                with open(os.path.join(config_dir, "config.txt"), "w") as f:
                    f.write(f"animation_path={self.animation_path}\n")
            except Exception as e:
                print(f"Error saving animation config: {e}")

    def select_files(self):
        """Select PDF files to process"""
        paths = filedialog.askopenfilenames(title="Choose PDF files", filetypes=[("PDF Files","*.pdf")])
        if not paths: 
            return
            
        self.file_paths = list(paths)
        self.files_list.delete(0, tk.END)
        
        for p in self.file_paths:
            self.files_list.insert(tk.END, os.path.basename(p))
            
        self.progress['value'] = 0
        self.detail_progress['value'] = 0
        self.status_label.config(text=f"{len(self.file_paths)} files selected")
        self.detail_label.config(text="")
        
        # Show pulse animation briefly
        self._show_loading_animation()
        self.after(1000, self._hide_loading_animation)

    def _show_loading_animation(self):
        """Show a loading animation"""
        if self.loading_animation:
            self.loading_animation.pack_forget()
            
        # Create a simple loading animation
        dots_frame = ttk.Frame(self.animation_frame)
        dots_frame.pack(pady=5)
        
        for i in range(5):
            dot = ttk.Label(dots_frame, text="●", font=("Arial", 12))
            dot.grid(row=0, column=i, padx=3)
            
            # Animated dot color change
            self.after(i * 150, lambda d=dot: d.config(foreground=self.theme["accent"]))
            self.after((i + 5) * 150, lambda d=dot: d.config(foreground=self.theme["fg"]))
        
        self.loading_animation = dots_frame

    def _hide_loading_animation(self):
        """Hide the loading animation"""
        if self.loading_animation:
            self.loading_animation.pack_forget()
            self.loading_animation = None

    def clear_files(self):
        """Clear selected files"""
        self.file_paths = []
        self.files_list.delete(0, tk.END)
        self.progress['value'] = 0
        self.detail_progress['value'] = 0
        self.status_label.config(text="File selection cleared")
        self.detail_label.config(text="")
        self._hide_loading_animation()

    def select_directory(self):
        """Select output directory"""
        d = filedialog.askdirectory(title="Select output folder")
        if not d: 
            return
            
        self.output_dir = d
        self.output_label.config(text=f"Output folder: {d}")

    def start_processing(self):
        """Start processing PDF files"""
        if not self.file_paths:
            messagebox.showwarning("Missing input", "Please select PDF files to process.")
            return
            
        if not self.output_dir:
            messagebox.showwarning("Missing output", "Please select an output folder.")
            return

        # Disable start button
        self.process_btn.config(state='disabled')
        self.progress['maximum'] = len(self.file_paths)
        self.progress['value'] = 0
        self.detail_progress['value'] = 0
        self.failures.clear()
        self.status_label.config(text="Starting batch processing...")
        self.detail_label.config(text="")
        
        # Show processing animation
        self._show_loading_animation()

        threading.Thread(target=self._run_batch, daemon=True).start()

    def _run_batch(self):
        """Process batch of PDFs in background thread"""
        for idx, pdf in enumerate(self.file_paths, 1):
            name = os.path.basename(pdf)
            out = os.path.join(self.output_dir, name)
            
            try:
                # Update file progress
                self.after(0, lambda n=name, i=idx: self._update_status_label(
                    f"Processing file {i}/{len(self.file_paths)}: {n}"))
                
                # Create processor with current settings
                processor = PDFProcessor(
                    pdf, 
                    out,
                    skip_first=self.skip_first_var.get(),
                    add_title=self.add_title_var.get(),
                    title_on_first_only=self.title_on_first_only_var.get(),
                    pages_per_sheet=self.pages_per_sheet_var.get()
                )
                
                # Process with page progress reporting
                processor.process(progress_callback=self._update_detail_progress)
                
            except Exception as e:
                self.failures.append((name, str(e)))
            finally:
                # Update file progress
                self.after(0, lambda i=idx: self._update_progress(i))
        
        self.after(0, self._finish)

    def _update_status_label(self, text):
        """Update status label from background thread"""
        self.status_label.config(text=text)
        
    def _update_progress(self, idx):
        """Update main progress bar from background thread"""
        self.progress['value'] = idx
        
    def _update_detail_progress(self, current, total):
        """Update detail progress bar from background thread"""
        self.after(0, lambda c=current, t=total: self._do_update_detail(c, t))
        
    def _do_update_detail(self, current, total):
        """Update detail progress UI elements"""
        self.detail_progress['maximum'] = total
        self.detail_progress['value'] = current
        self.detail_label.config(text=f"Processing page {current}/{total}")
        
        # Periodic animation update to show activity
        if current % 3 == 0:
            self._refresh_loading_animation()
            
    def _refresh_loading_animation(self):
        """Refresh the loading animation during processing"""
        if self.loading_animation:
            self.loading_animation.pack_forget()
        
        self._show_loading_animation()

    def _finish(self):
        """Clean up after processing completes"""
        self.process_btn.config(state='normal')
        self._hide_loading_animation()
        
        if self.failures:
            msg = "Errors in processing:\n" + "\n".join(f"{n}: {err}" for n, err in self.failures)
            messagebox.showerror("Batch Completed with Errors", msg)
            self.status_label.config(text="Completed with errors", foreground=self.theme["status_error"])
        else:
            messagebox.showinfo("Batch Completed", "All files processed successfully")
            self.status_label.config(text="All done!", foreground=self.theme["status_good"])
            
            # Show success animation
            self._show_success_animation()
            
        self.detail_progress['value'] = 0
        self.detail_label.config(text="")
        
    def _show_success_animation(self):
        """Show success animation"""
        success_frame = ttk.Frame(self.animation_frame)
        success_frame.pack(pady=5)
        
        # Create checkmark animation
        checkmark = ttk.Label(success_frame, text="✓", font=("Arial", 24),
                             foreground=self.theme["status_good"])
        checkmark.pack()
        
        # Animate checkmark
        def _animate_checkmark(size=24, direction=1):
            if 20 <= size <= 32:
                checkmark.config(font=("Arial", size))
                self.after(50, lambda: _animate_checkmark(size + direction, 
                                                         direction if size < 32 and direction > 0 
                                                         else -1 if size >= 32 
                                                         else -direction if size <= 20 else direction))
            else:
                # End animation after a few cycles
                self.after(1000, lambda: success_frame.pack_forget())
                
        _animate_checkmark()
        
        self.loading_animation = success_frame

# Load saved configuration
def load_config():
    """Load saved configuration"""
    config = {}
    try:
        config_path = os.path.join(os.path.expanduser("~"), ".pdf_processor", "config.txt")
        if os.path.exists(config_path):
            with open(config_path, "r") as f:
                for line in f:
                    if "=" in line:
                        key, value = line.strip().split("=", 1)
                        config[key] = value
    except Exception as e:
        print(f"Error loading config: {e}")
        
    return config

if __name__ == "__main__":
    # Load config
    config = load_config()
    
    # Start app
    app = App()
    
    # Set animation path from config if it exists
    if "animation_path" in config and os.path.exists(config["animation_path"]):
        app.animation_path = config["animation_path"]
    
    # Apply some styling
    app.mainloop()
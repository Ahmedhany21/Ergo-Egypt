
## THIS Projects is doing the same job of excel power query function 'Unpivoting columns' but it was making for uploading multiple sheets and manipulating each sheet on it's own by doing some work on it
#  1. Adding files -> Adding file used to open your own pc to upload your sheets 
#  2. Clear all files -> Used to  clear all files you uploaded to put in new ones
#  3. Process next file -> Used to preview your current file to see the issues in text box preview 
#  4. Select Columns -> Used to select the columns you want to unpivot on it 
#  5. CLean Data -> Used to clear all 'NAN', 'blank' and Empty values after selecting you columns
#  6. flip layout -> After previewing your Excel Sheet if the layout is in from left to right ot right to left you can make it using this button
#  7. Finish Current File -> Used to finish manipulating on your selected file and going through the next one
#  8. UNdo -> To undo your unwanted action
#  9. Save all results  -> to put all files you already finished working on it to be one single file ready to use in Excel sheet
#  10. Exit -> This button closes the Program


import customtkinter as ctk
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import numpy as np
import os
from datetime import datetime
from PIL import Image, ImageTk
import sys
import ctypes

# Set appearance mode and default color theme
ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

# Light blue color scheme
LIGHT_BLUE = "#FFFFFF"
PRIMARY_BLUE = "#4DA6FF"
DARK_BLUE = "#0066CC"
ACCENT_BLUE = "#80CCFF"
WHITE = "#FFFFFF"

class MultiFileUnpivotApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        
        # Configure window
        self.title("Multi-File Excel/CSV Unpivot Tool")
        self.geometry("1200x750")
        
        # ========== CUSTOM ICON CODE ==========
        try:
            # For Windows taskbar icon
            if sys.platform == "win32":
                myappid = 'yourcompany.multifileunpivottool.1.0'
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
            
            # Load icon using PIL
            # FIX 1: Use raw string or double backslashes for Windows paths
            icon_path = r"C:\Users\600 G3\Downloads\Phone-Rotate-2--Streamline-Ultimate.ico"
            
            # FIX 2: Alternative path format with double backslashes
            # icon_path = "C:\\Users\\600 G3\\Downloads\\Phone-Rotate-2--Streamline-Ultimate.png"
            
            # FIX 3: Use relative path if icon is in same folder as script
            # icon_path = "app_icon.png"  # Put icon file in same folder
            
            # Load and set the icon
            icon_image = Image.open(icon_path)
            
            # For Tkinter compatibility
            photo = ImageTk.PhotoImage(icon_image)
            
            # Set window icon (works for title bar)
            self.wm_iconphoto(False, photo)
            
            # Keep reference to prevent garbage collection
            self.icon_image = photo
            
            print(f"Icon loaded successfully from: {icon_path}")
            
        except FileNotFoundError:
            print("Icon file not found. Using default icon.")
            # You can create a simple default icon here
            self.create_default_icon()
        except Exception as e:
            print(f"Could not set icon: {e}")
            self.create_default_icon()
        # =======================================
        
        # Configure grid layout
        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(0, weight=1)
        
        # Application state
        self.files_to_process = []  # List of file paths
        self.current_file_index = 0  # Index of current file being processed
        self.df = None  # Current dataframe
        self.processed_dataframes = []  # Store processed dataframes
        self.df_history = []  # Store history for undo
        self.current_step = -1
        self.selected_columns = []
        
        # Dictionary to store file list labels for hover functionality
        self.file_labels = {}
        
        # Create sidebar frame
        self.sidebar_frame = ctk.CTkFrame(self, width=220, corner_radius=0, fg_color=PRIMARY_BLUE)
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(10, weight=1)
        
        # Logo label
        self.logo_label = ctk.CTkLabel(
            self.sidebar_frame, 
            text="Multi-File\nUnpivot Tool", 
            font=ctk.CTkFont(size=20, weight="bold"),
            justify="center"
        )
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        
        # Sidebar buttons
        self.add_files_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="Add Files", 
            command=self.add_files,
            fg_color=DARK_BLUE,
            hover_color=ACCENT_BLUE
        )
        self.add_files_btn.grid(row=1, column=0, padx=20, pady=10)
        
        self.clear_files_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="Clear All Files", 
            command=self.clear_all_files,
            fg_color="#FF3333",
            hover_color="#FF6666"
        )
        self.clear_files_btn.grid(row=2, column=0, padx=20, pady=5)
        
        self.process_next_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="Process Next File", 
            command=self.process_next_file,
            fg_color=DARK_BLUE,
            hover_color=ACCENT_BLUE,
            state="disabled"
        )
        self.process_next_btn.grid(row=3, column=0, padx=20, pady=10)
        
        self.unpivot_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="Select Columns", 
            command=self.select_columns_window,
            fg_color=DARK_BLUE,
            hover_color=ACCENT_BLUE,
            state="disabled"
        )
        self.unpivot_btn.grid(row=4, column=0, padx=20, pady=10)
        
        # Clean data button
        self.clean_data_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="Clean Data", 
            command=self.clean_data,
            fg_color="#505050",
            hover_color="#BDB1A0",
            state="disabled"
        )
        self.clean_data_btn.grid(row=5, column=0, padx=20, pady=10)
        
        # Flip Layout button
        self.flip_layout_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="🔄 Flip Layout", 
            command=self.flip_current_layout,
            fg_color="#FF9900",
            hover_color="#FFCC00",
            state="disabled"
        )
        self.flip_layout_btn.grid(row=6, column=0, padx=20, pady=10)
        
        self.finish_file_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="Finish Current File", 
            command=self.finish_current_file,
            fg_color="#00AA00",
            hover_color="#00CC00",
            state="disabled"
        )
        self.finish_file_btn.grid(row=7, column=0, padx=20, pady=10)
        
        # Undo button
        self.undo_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="↩ Undo", 
            command=self.undo_action,
            fg_color="#505050",
            hover_color="#BDB1A0",
            state="disabled"
        )
        self.undo_btn.grid(row=8, column=0, padx=20, pady=10)
        
        # Save all button
        self.save_all_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="Save All Results", 
            command=self.save_all_results,
            fg_color="#008F18",
            hover_color="#6BFF6B",
            state="disabled"
        )
        self.save_all_btn.grid(row=9, column=0, padx=20, pady=10)
        
        # Exit button
        self.exit_btn = ctk.CTkButton(
            self.sidebar_frame, 
            text="Exit", 
            command=self.quit,
            fg_color="#FF3333",
            hover_color="#FF6666"
        )
        self.exit_btn.grid(row=10, column=0, padx=20, pady=(10, 20))
        
        # Create main content area
        self.main_frame = ctk.CTkFrame(self, fg_color="transparent")
        self.main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        self.main_frame.grid_columnconfigure(0, weight=1)
        self.main_frame.grid_rowconfigure(2, weight=1)
        
        # Title label
        self.title_label = ctk.CTkLabel(
            self.main_frame, 
            text="Add multiple files to begin batch unpivoting",
            font=ctk.CTkFont(size=24, weight="bold"),
            text_color="black"
        )
        self.title_label.grid(row=0, column=0, padx=20, pady=(20, 10), sticky="w")
        
        # File list frame with scrollable functionality
        self.file_list_frame = ctk.CTkFrame(self.main_frame, fg_color=WHITE)
        self.file_list_frame.grid(row=1, column=0, padx=20, pady=(10, 10), sticky="ew")
        self.file_list_frame.grid_columnconfigure(0, weight=1)
        
        # Create a scrollable frame for file list
        self.file_list_scrollable = ctk.CTkScrollableFrame(
            self.file_list_frame, 
            fg_color=WHITE,
            height=150
        )
        self.file_list_scrollable.grid(row=0, column=0, padx=10, pady=10, sticky="nsew")
        self.file_list_scrollable.grid_columnconfigure(0, weight=1)
        
        self.file_list_title = ctk.CTkLabel(
            self.file_list_scrollable,
            text="No files added yet",
            font=ctk.CTkFont(size=14),
            text_color="black"
        )
        self.file_list_title.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        # Current file info frame - Positioned between file list and preview
        self.current_file_frame = ctk.CTkFrame(self.main_frame, fg_color=WHITE)
        self.current_file_frame.grid(row=2, column=0, padx=20, pady=(10, 10), sticky="ew")
        self.current_file_frame.grid_columnconfigure(0, weight=1)
        
        # Current file info and remove button in one row
        self.current_file_info_frame = ctk.CTkFrame(self.current_file_frame, fg_color="transparent")
        self.current_file_info_frame.grid(row=0, column=0, padx=5, pady=5, sticky="ew")
        self.current_file_info_frame.grid_columnconfigure(0, weight=1)
        self.current_file_info_frame.grid_columnconfigure(1, weight=0)
        
        self.current_file_label = ctk.CTkLabel(
            self.current_file_info_frame, 
            text="No file currently being processed",
            font=ctk.CTkFont(size=14, weight="bold"),
            text_color=DARK_BLUE
        )
        self.current_file_label.grid(row=0, column=0, padx=5, pady=5, sticky="w")
        
        # Remove Current File button - positioned next to current file label
        self.remove_current_btn = ctk.CTkButton(
            self.current_file_info_frame,
            text="✕ Remove This File",
            command=self.remove_current_file,
            fg_color="#FF3333",
            hover_color="#FF6666",
            width=120,
            state="disabled"
        )
        self.remove_current_btn.grid(row=0, column=1, padx=(20, 5), pady=5, sticky="e")
        
        # Create textbox for data preview - Below the current file info
        self.preview_text = ctk.CTkTextbox(self.main_frame, width=500)
        self.preview_text.grid(row=3, column=0, padx=20, pady=(10, 20), sticky="nsew")
        self.preview_text.insert("0.0", "Data preview will appear here when processing files.")
        self.preview_text.configure(state="disabled")
        
        # Status bar
        self.status_bar = ctk.CTkLabel(
            self, 
            text="Ready to add files",
            font=ctk.CTkFont(size=12),
            text_color=DARK_BLUE,
            anchor="w"
        )
        self.status_bar.grid(row=1, column=0, columnspan=2, sticky="ew", padx=20, pady=(0, 10))
        
        # Progress label
        self.progress_label = ctk.CTkLabel(
            self.main_frame,
            text="",
            font=ctk.CTkFont(size=12),
            text_color=DARK_BLUE
        )
        self.progress_label.grid(row=4, column=0, padx=20, pady=(0, 10), sticky="w")
    
    def create_default_icon(self):
        """Create a simple default icon if custom icon fails"""
        try:
            # Create a simple blue square as default icon
            img = Image.new('RGB', (64, 64), color=PRIMARY_BLUE)
            photo = ImageTk.PhotoImage(img)
            self.wm_iconphoto(False, photo)
            self.icon_image = photo  # Keep reference
            print("Using default blue icon")
        except:
            pass  # If all fails, just use no icon
    
    def update_status(self, message):
        """Update status bar message"""
        self.status_bar.configure(text=f"Status: {message}")
        self.update_idletasks()
    
    def create_tooltip(self, widget, text):
        """Create a simple tooltip for a widget"""
        def enter(event):
            widget.configure(text=text, cursor="hand2")
        
        def leave(event):
            widget.configure(text=widget.original_text, cursor="")
        
        # Store the original text
        widget.original_text = widget.cget("text")
        
        # Bind events
        widget.bind("<Enter>", enter)
        widget.bind("<Leave>", leave)
    
    def update_file_list_display(self):
        """Update the file list display with hover functionality"""
        # Clear existing file labels
        for label in self.file_labels.values():
            label.destroy()
        self.file_labels.clear()
        
        if self.files_to_process:
            self.file_list_title.configure(text=f"Files to process ({len(self.files_to_process)}):")
            
            for i, file_path in enumerate(self.files_to_process):
                filename = os.path.basename(file_path)
                
                # Determine status symbol and color
                if i < self.current_file_index:
                    status = "✓"
                    color = "#00AA00"  # Green for processed
                elif i == self.current_file_index:
                    status = "●"
                    color = "#FF9900"  # Orange for current
                else:
                    status = "○"
                    color = "#666666"  # Gray for pending
                
                # Create label for each file
                file_label = ctk.CTkLabel(
                    self.file_list_scrollable,
                    text=f"{status} {filename}",
                    font=ctk.CTkFont(size=12),
                    text_color=color,
                    anchor="w"
                )
                file_label.grid(row=i+1, column=0, padx=5, pady=2, sticky="w")
                
                # Store for later reference
                self.file_labels[i] = file_label
                
                # Add hover functionality to show full path
                self.create_tooltip(file_label, file_path)
        else:
            self.file_list_title.configure(text="No files added yet")
    
    def update_progress_label(self):
        """Update progress label"""
        if self.files_to_process:
            processed = len(self.processed_dataframes)
            total = len(self.files_to_process)
            remaining = total - self.current_file_index
            self.progress_label.configure(
                text=f"Progress: {processed} processed, {remaining} remaining"
            )
        else:
            self.progress_label.configure(text="")
    
    def save_state(self):
        """Save current dataframe state for undo functionality"""
        if self.df is not None:
            # Store a copy of the dataframe
            self.df_history = self.df_history[:self.current_step + 1]  # Trim any forward history
            self.df_history.append(self.df.copy())
            self.current_step = len(self.df_history) - 1
            
            # Enable undo button if we have at least one state
            if len(self.df_history) > 0:
                self.undo_btn.configure(state="normal")
    
    def undo_action(self):
        """Revert to previous state for current file"""
        if len(self.df_history) > 0 and self.current_step > 0:
            self.current_step -= 1
            self.df = self.df_history[self.current_step].copy()
            self.update_preview()
            self.update_status(f"Undo completed. Reverted to previous state.")
            
            # Update button states based on current state
            self.update_current_file_buttons()
            
        elif self.current_step == 0 and len(self.df_history) > 0:
            # We're at the first state - revert to originally loaded file
            self.df = self.df_history[0].copy()
            self.current_step = 0
            self.update_preview()
            self.update_status("Reverted to original file state (as first loaded).")
    
    def update_current_file_buttons(self):
        """Update button states for current file processing"""
        if self.df is not None:
            self.unpivot_btn.configure(state="normal")
            self.clean_data_btn.configure(state="normal")
            self.flip_layout_btn.configure(state="normal")
            self.finish_file_btn.configure(state="normal")
            self.remove_current_btn.configure(state="normal")
        else:
            self.unpivot_btn.configure(state="disabled")
            self.clean_data_btn.configure(state="disabled")
            self.flip_layout_btn.configure(state="disabled")
            self.finish_file_btn.configure(state="disabled")
            self.remove_current_btn.configure(state="disabled")
    
    def flip_current_layout(self):
        """Flip the current dataframe page layout horizontally (RTL to LTR)"""
        if self.df is not None:
            # Save state before flipping
            self.save_state()
            
            # Get column names before flipping (for display)
            before_columns = list(self.df.columns)
            
            # Flip the dataframe horizontally - this is page layout flip!
            # Reverse the order of all columns
            self.df = self.df.iloc[:, ::-1]
            
            # Get column names after flipping (for display)
            after_columns = list(self.df.columns)
            
            # Update UI
            self.update_preview()
            self.update_status("Page layout flipped (mirrored horizontally). Undo available.")
            
            # Show visual confirmation
            confirmation_msg = (
                "✅ PAGE LAYOUT FLIPPED\n\n"
                "Right-to-Left (RTL) → Left-to-Right (LTR)\n\n"
                "What changed:\n"
                "• First column moved to last position\n"
                "• Last column moved to first position\n"
                "• All columns reversed order\n\n"
                "Example (simplified):\n"
                "Before: [ID, Name, Date, Value1, Value2, Value3]\n"
                "After:  [Value3, Value2, Value1, Date, Name, ID]\n\n"
                "This fixes Arabic Excel files that open with\n"
                "columns in reverse order due to RTL settings."
            )
            
            messagebox.showinfo("Page Layout Flipped", confirmation_msg)
    
    def add_files(self):
        """Add multiple Excel or CSV files"""
        filetypes = [
            ("Excel files", "*.xlsx *.xls"),
            ("CSV files", "*.csv"),
            ("All files", "*.*")
        ]
        
        file_paths = filedialog.askopenfilenames(
            title="Select files to process",
            filetypes=filetypes
        )
        
        if file_paths:
            # Add new files to the list
            new_files = [path for path in file_paths if path not in self.files_to_process]
            self.files_to_process.extend(new_files)
            
            self.update_file_list_display()  # Updated to use new display method
            self.update_progress_label()
            
            if len(self.files_to_process) > 0:
                self.process_next_btn.configure(state="normal")
                self.update_status(f"Added {len(new_files)} file(s). Total: {len(self.files_to_process)} files ready to process.")
            else:
                self.update_status("No new files added.")
    
    def clear_all_files(self):
        """Clear all files from the list"""
        if self.files_to_process:
            if messagebox.askyesno("Clear Files", "Are you sure you want to clear all files?"):
                self.files_to_process = []
                self.processed_dataframes = []
                self.current_file_index = 0
                self.df = None
                
                # Reset UI
                self.update_file_list_display()  # Updated to use new display method
                self.update_progress_label()
                self.current_file_label.configure(text="No file currently being processed")
                self.preview_text.configure(state="normal")
                self.preview_text.delete("0.0", "end")
                self.preview_text.insert("0.0", "Data preview will appear here when processing files.")
                self.preview_text.configure(state="disabled")
                
                # Disable buttons
                self.process_next_btn.configure(state="disabled")
                self.unpivot_btn.configure(state="disabled")
                self.clean_data_btn.configure(state="disabled")
                self.flip_layout_btn.configure(state="disabled")
                self.finish_file_btn.configure(state="disabled")
                self.save_all_btn.configure(state="disabled")
                self.undo_btn.configure(state="disabled")
                self.remove_current_btn.configure(state="disabled")
                
                self.update_status("All files cleared. Ready to add new files.")
    
    def process_next_file(self):
        """Load and process the next file in the list"""
        if self.current_file_index < len(self.files_to_process):
            file_path = self.files_to_process[self.current_file_index]
            filename = os.path.basename(file_path)
            
            self.current_file_label.configure(text=f"Processing: {filename}")
            self.remove_current_btn.configure(state="normal")
            self.update_status(f"Loading {filename}...")
            
            try:
                # Read file with default settings (pandas will use first row as headers)
                if file_path.endswith('.csv'):
                    self.df = pd.read_csv(file_path)
                else:  # Excel file
                    self.df = pd.read_excel(file_path)
                
                # Check if file has data
                if len(self.df) < 1:
                    response = messagebox.askquestion(
                        "Empty File",
                        f"File '{filename}' appears to have no data.\n\n"
                        f"Rows in file: {len(self.df)}\n\n"
                        "What would you like to do?",
                        type=messagebox.YESNOCANCEL,
                        icon='warning'
                    )
                    
                    if response == 'yes':  # Skip
                        self.skip_current_file()
                        return
                    elif response == 'cancel':  # Remove
                        self.remove_current_file()
                        return
                
                # Clear history and save initial state
                self.df_history = [self.df.copy()]
                self.current_step = 0
                
                # Update UI
                self.update_preview()
                self.update_current_file_buttons()
                self.undo_btn.configure(state="normal")
                self.update_status(f"File loaded. Shape: {self.df.shape}")
                
            except Exception as e:
                # File loading error
                error_msg = str(e)
                response = messagebox.askquestion(
                    "File Loading Error",
                    f"Failed to load file '{filename}':\n\n{error_msg}\n\n"
                    "Do you want to remove this file from the list?",
                    icon='error'
                )
                
                if response == 'yes':
                    self.remove_current_file()
                else:
                    # Just skip to next
                    self.skip_current_file()
        else:
            self.update_status("No more files to process.")

    def skip_current_file(self):
        """Skip the current file without processing it (internal function)"""
        if self.current_file_index < len(self.files_to_process):
            filename = os.path.basename(self.files_to_process[self.current_file_index])
            
            # Move to next file without saving current file
            self.current_file_index += 1
            
            # Reset current file state
            self.df = None
            self.df_history = []
            self.current_step = -1
            
            # Update UI
            self.update_file_list_display()
            self.update_progress_label()
            self.undo_btn.configure(state="disabled")
            self.remove_current_btn.configure(state="disabled")
            
            if self.current_file_index < len(self.files_to_process):
                # Still have files to process
                next_filename = os.path.basename(self.files_to_process[self.current_file_index])
                self.current_file_label.configure(text=f"File skipped. Next: {next_filename}")
                self.preview_text.configure(state="normal")
                self.preview_text.delete("0.0", "end")
                self.preview_text.insert("0.0", f"File skipped. {len(self.files_to_process) - self.current_file_index} file(s) remaining.\nClick 'Process Next File' to continue.")
                self.preview_text.configure(state="disabled")
                
                # Update button states
                self.unpivot_btn.configure(state="disabled")
                self.clean_data_btn.configure(state="disabled")
                self.flip_layout_btn.configure(state="disabled")
                self.finish_file_btn.configure(state="disabled")
                
                self.update_status(f"File skipped. {len(self.files_to_process) - self.current_file_index} file(s) remaining.")
            else:
                # All files processed or skipped
                self.handle_all_files_completed()
        else:
            self.update_status("No more files to process.")
    
    def remove_current_file(self):
        """Remove the current file from the list"""
        if self.current_file_index < len(self.files_to_process):
            filename = os.path.basename(self.files_to_process[self.current_file_index])
            
            # Remove current file from list
            removed_file = self.files_to_process.pop(self.current_file_index)
            
            # Don't increment current_file_index since we removed the current file
            # The next file in list becomes the new current file
            
            # Reset current file state
            self.df = None
            self.df_history = []
            self.current_step = -1
            
            # Update UI
            self.update_file_list_display()
            self.update_progress_label()
            self.undo_btn.configure(state="disabled")
            self.remove_current_btn.configure(state="disabled")
            
            if self.files_to_process:
                # Still have files to process
                if self.current_file_index < len(self.files_to_process):
                    # Show next file
                    next_filename = os.path.basename(self.files_to_process[self.current_file_index])
                    self.current_file_label.configure(text=f"File removed. Next: {next_filename}")
                else:
                    # We removed the last file
                    self.current_file_index = 0
                    self.current_file_label.configure(text="No file currently being processed")
                
                self.preview_text.configure(state="normal")
                self.preview_text.delete("0.0", "end")
                self.preview_text.insert("0.0", f"File '{filename}' removed from list.\n{len(self.files_to_process)} file(s) remaining.\nClick 'Process Next File' to continue.")
                self.preview_text.configure(state="disabled")
                
                # Update button states
                self.unpivot_btn.configure(state="disabled")
                self.clean_data_btn.configure(state="disabled")
                self.flip_layout_btn.configure(state="disabled")
                self.finish_file_btn.configure(state="disabled")
                
                self.update_status(f"File '{filename}' removed. {len(self.files_to_process)} file(s) remaining.")
            else:
                # No files left
                self.current_file_label.configure(text="No file currently being processed")
                self.preview_text.configure(state="normal")
                self.preview_text.delete("0.0", "end")
                self.preview_text.insert("0.0", f"File '{filename}' removed. No files remaining.")
                self.preview_text.configure(state="disabled")
                
                # Disable buttons
                self.process_next_btn.configure(state="disabled")
                self.unpivot_btn.configure(state="disabled")
                self.clean_data_btn.configure(state="disabled")
                self.flip_layout_btn.configure(state="disabled")
                self.finish_file_btn.configure(state="disabled")
                self.save_all_btn.configure(state="disabled")
                
                self.update_status(f"File '{filename}' removed. No files remaining.")
        else:
            self.update_status("No file to remove.")
    
    def update_preview(self):
        """Update the preview textbox with dataframe content"""
        self.preview_text.configure(state="normal")
        self.preview_text.delete("0.0", "end")
        
        if self.df is not None:
            # Show first 100 rows
            preview_df = self.df.head(100)
            
            # Convert dataframe to string
            df_string = preview_df.to_string(max_rows=100, max_cols=20)
            
            # Insert into textbox
            current_file = os.path.basename(self.files_to_process[self.current_file_index]) if self.current_file_index < len(self.files_to_process) else "Unknown"
            self.preview_text.insert("0.0", f"File: {current_file}\n")
            self.preview_text.insert("end", f"Data Preview ({self.df.shape[0]} rows, {self.df.shape[1]} columns):\n")
            self.preview_text.insert("end", "="*80 + "\n")
            
            # Show column names
            self.preview_text.insert("end", "Columns:\n")
            for i, col in enumerate(self.df.columns):
                self.preview_text.insert("end", f"  {i+1}. {col}\n")
            
            self.preview_text.insert("end", "\n" + "="*80 + "\n")
            self.preview_text.insert("end", "First few rows:\n")
            self.preview_text.insert("end", df_string)
        
        self.preview_text.configure(state="disabled")
    
    def handle_all_files_completed(self):
        """Handle UI when all files are processed"""
        self.current_file_label.configure(text="All files processed!")
        self.preview_text.configure(state="normal")
        self.preview_text.delete("0.0", "end")
        
        # Show summary of processed files
        summary_text = "All files have been processed!\n\n"
        summary_text += f"Total files processed: {len(self.processed_dataframes)}\n"
        summary_text += f"Total files in original list: {len(self.files_to_process) + len(self.processed_dataframes)}\n\n"
        summary_text += "Summary of processed files:\n"
        for i, item in enumerate(self.processed_dataframes):
            df = item['dataframe']
            summary_text += f"{i+1}. {item['filename']}: {df.shape[0]} rows, {df.shape[1]} columns\n"
        
        self.preview_text.insert("0.0", summary_text)
        self.preview_text.configure(state="disabled")
        
        # Update button states
        self.process_next_btn.configure(state="disabled")
        self.unpivot_btn.configure(state="disabled")
        self.clean_data_btn.configure(state="disabled")
        self.flip_layout_btn.configure(state="disabled")
        self.finish_file_btn.configure(state="disabled")
        self.remove_current_btn.configure(state="disabled")
        
        if self.processed_dataframes:
            self.save_all_btn.configure(state="normal")
            self.update_status(f"All {len(self.processed_dataframes)} files processed successfully. Ready to save combined results.")
        else:
            self.save_all_btn.configure(state="disabled")
            self.update_status("All files processed or skipped. No data to save.")
    
    def clean_data(self):
        """Clean the data by removing empty rows and footer rows"""
        if self.df is None:
            return
        
        # Save current state before cleaning
        self.save_state()
        
        original_rows = len(self.df)
        
        # Remove completely empty rows (all NaN)
        self.df = self.df.dropna(how='all')
        
        # Remove rows where all ID columns (first few columns) are empty
        # Common ID columns: ID, Date, Name, etc.
        id_columns = []
        for col in self.df.columns:
            col_str = str(col).lower()
            if any(keyword in col_str for keyword in ['id', 'date', 'name', 'employee', 'code']):
                id_columns.append(col)
        
        if id_columns:
            # Keep rows where at least one ID column has data
            mask = self.df[id_columns].isna().all(axis=1)
            self.df = self.df[~mask]
        
        # Remove rows with common footer text in any column
        footer_keywords = [
            'أجازة', 'غياب', 'جمعة',  # Arabic keywords from your example
            'vacation', 'absent', 'friday',  # English equivalents
            'إجازة', 'تقرير', 'ملاحظات'  # Other possible Arabic footers
        ]
        
        # Convert all columns to string for searching
        for col in self.df.columns:
            self.df[col] = self.df[col].astype(str)
        
        # Create a mask to identify footer rows
        footer_mask = pd.Series(False, index=self.df.index)
        for keyword in footer_keywords:
            # Check if keyword appears in any cell
            for col in self.df.columns:
                if col not in id_columns:  # Don't check ID columns for footer text
                    contains_keyword = self.df[col].str.contains(keyword, case=False, na=False)
                    footer_mask = footer_mask | contains_keyword
        
        # Remove footer rows
        self.df = self.df[~footer_mask]
        
        # Convert back from string to appropriate types
        for col in self.df.columns:
            # Try to convert back to numeric where possible
            try:
                self.df[col] = pd.to_numeric(self.df[col], errors='ignore')
            except:
                pass
        
        # Reset index after cleaning
        self.df = self.df.reset_index(drop=True)
        
        removed_rows = original_rows - len(self.df)
        
        # Save state after cleaning
        self.save_state()
        
        # Update UI
        self.update_preview()
        self.update_status(f"Cleaned data: Removed {removed_rows} empty/footer rows. New shape: {self.df.shape}")
        
        # Show summary
        # In the clean_data() method, update the messagebox at the end:
        messagebox.showinfo("Data Cleaned", 
                        f"Removed {removed_rows} data rows:\n"
                        f"- Completely empty rows\n"
                        f"- Rows with missing ID information\n"
                        f"- Rows with footer text (أجازة أو غياب, يوم جمعة, etc.)\n\n"
                        f"Original data rows: {original_rows}\n"
                        f"After cleaning: {len(self.df)} data rows\n\n"
                        f"Note: Column headers are not counted as data rows.")
    
    def select_columns_window(self):
        """Open window to select columns for unpivoting"""
        if self.df is None:
            messagebox.showwarning("No Data", "Please process a file first.")
            return
        
        # First, read the file RAW to see the actual Excel structure
        file_path = self.files_to_process[self.current_file_index]
        filename = os.path.basename(file_path)
        
        try:
            # Read file WITHOUT headers to see actual structure
            if file_path.endswith('.csv'):
                df_raw = pd.read_csv(file_path, header=None)
            else:
                df_raw = pd.read_excel(file_path, header=None)
            
            # Check if we have at least 2 rows
            if len(df_raw) >= 2:
                excel_row_1 = df_raw.iloc[0]  # Actual Excel Row 1
                excel_row_2 = df_raw.iloc[1]  # Actual Excel Row 2
                
                # Show what's in Excel Row 1
                row1_preview = []
                for i, val in enumerate(excel_row_1.head(6)):  # Show first 6 columns
                    if pd.notna(val):
                        row1_preview.append(f"Col {i+1}: '{val}'")
                
                row1_text = "\n".join(row1_preview) if row1_preview else "(All cells are empty)"
                
                # Show what's in Excel Row 2 (current headers)
                row2_preview = []
                for i, val in enumerate(excel_row_2.head(6)):  # Show first 6 columns
                    if pd.notna(val):
                        row2_preview.append(f"Col {i+1}: '{val}'")
                
                row2_text = "\n".join(row2_preview) if row2_preview else "(No headers)"
                
                response = messagebox.askyesnocancel(
                    "Excel File Structure",
                    f"File: {filename}\n\n"
                    f"ACTUAL EXCEL STRUCTURE:\n"
                    f"═══════════════════════════\n"
                    f"EXCEL ROW 1 (currently loaded as headers):\n{row1_text}\n\n"
                    f"EXCEL ROW 2 (actual column headers):\n{row2_text}\n\n"
                    f"EXCEL ROW 3+ (data records)\n\n"
                    f"═══════════════════════════\n"
                    f"Do you want to FIX THE HEADERS?\n\n"
                    f"• YES: Remove Excel Row 1, use Excel Row 2 as headers\n"
                    f"      ↳ Excel Row 1 will be discarded\n"
                    f"      ↳ Excel Row 2 becomes column headers\n"
                    f"      ↳ Excel Row 3+ becomes data\n\n"
                    f"• NO: Keep current structure (Excel Row 1 as headers)\n"
                    f"• CANCEL: Go back to main window"
                )
                
                if response is None:  # Cancel
                    self.update_status("Column selection cancelled.")
                    return
                elif response:  # Yes - Fix headers
                    # Save state before fixing
                    self.save_state()
                    
                    # Use Excel Row 2 as headers
                    df_raw.columns = excel_row_2.tolist()
                    # Remove first 2 rows (Excel Row 1 and Row 2)
                    self.df = df_raw.iloc[2:].reset_index(drop=True)
                    
                    # Clean column names
                    self.df.columns = [str(col) if pd.notna(col) else f"Column_{i}" for i, col in enumerate(self.df.columns)]
                    
                    # Update UI
                    self.update_preview()
                    self.update_status(f"Headers fixed! Used Excel Row 2 as headers. Data rows: {len(self.df)}")
                    
                    # Show confirmation
                    messagebox.showinfo("Headers Fixed",
                                    f"Successfully fixed headers!\n\n"
                                    f"• Removed: Excel Row 1 (contained day names/empty)\n"
                                    f"• New headers: Excel Row 2 values\n"
                                    f"• Data starts from: Excel Row 3\n\n"
                                    f"New data shape: {self.df.shape}\n"
                                    f"You can use UNDO if this was incorrect.")
            
            else:
                # File doesn't have enough rows for this structure
                messagebox.showinfo("File Structure",
                                f"File '{filename}' has {len(df_raw)} rows.\n"
                                f"Cannot detect the header structure you described.\n"
                                f"Proceeding with current data.")
        
        except Exception as e:
            messagebox.showerror("Error", f"Could not read file structure: {str(e)}")
            self.update_status("Error reading file structure")
        
        # Now continue with column selection window
        self.open_column_selection_window()

    def open_column_selection_window(self):
        """Open the actual column selection window"""
        # Create new window
        self.select_window = ctk.CTkToplevel(self)
        self.select_window.title("Select Columns for Unpivoting")
        self.select_window.geometry("700x600")  # Slightly larger for new buttons
        self.select_window.grab_set()  # Make window modal
        
        # Configure grid
        self.select_window.grid_columnconfigure(0, weight=1)
        self.select_window.grid_rowconfigure(2, weight=1)  # Changed to row 2 for selection buttons
        
        # Title
        current_file = os.path.basename(self.files_to_process[self.current_file_index])
        title_label = ctk.CTkLabel(
            self.select_window, 
            text=f"Select columns to unpivot for: {current_file}\n(ID columns that will remain as rows)",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color=DARK_BLUE
        )
        title_label.grid(row=0, column=0, padx=20, pady=20, sticky="w")
        
        # Info label
        info_label = ctk.CTkLabel(
            self.select_window,
            text="✓ Selected columns will be kept as ID columns\n⚫ Unselected columns will be unpivoted",
            font=ctk.CTkFont(size=12),
            text_color=DARK_BLUE
        )
        info_label.grid(row=0, column=0, padx=20, pady=(80, 5), sticky="w")
        
        # Quick selection buttons frame
        selection_buttons_frame = ctk.CTkFrame(self.select_window, fg_color="transparent")
        selection_buttons_frame.grid(row=1, column=0, padx=20, pady=(0, 10), sticky="ew")
        selection_buttons_frame.grid_columnconfigure(0, weight=1)
        selection_buttons_frame.grid_columnconfigure(1, weight=1)
        
        # Select All button
        select_all_btn = ctk.CTkButton(
            selection_buttons_frame,
            text="Select All Columns",
            command=self.select_all_columns,
            fg_color="#00AA00",
            hover_color="#00CC00",
            width=150
        )
        select_all_btn.grid(row=0, column=0, padx=5, pady=5)
        
        # Deselect All button
        deselect_all_btn = ctk.CTkButton(
            selection_buttons_frame,
            text="Deselect All Columns",
            command=self.deselect_all_columns,
            fg_color="#FF3333",
            hover_color="#FF6666",
            width=150
        )
        deselect_all_btn.grid(row=0, column=1, padx=5, pady=5)
        
        # Create frame for checkboxes
        checkbox_frame = ctk.CTkScrollableFrame(self.select_window, fg_color=LIGHT_BLUE)
        checkbox_frame.grid(row=2, column=0, padx=20, pady=(0, 20), sticky="nsew")
        
        # Create checkboxes for each column
        self.column_vars = {}
        for i, col in enumerate(self.df.columns):
            var = tk.BooleanVar()
            # Pre-select common ID columns (case-insensitive)
            col_str = str(col).lower()
            if any(keyword in col_str for keyword in ['id', 'date', 'name', 'location', 'job', 'color']):
                var.set(True)
            
            chk = ctk.CTkCheckBox(
                checkbox_frame, 
                text=str(col),
                variable=var,
                onvalue=True,
                offvalue=False
            )
            chk.grid(row=i, column=0, padx=10, pady=5, sticky="w")
            self.column_vars[col] = var
        
        # Action buttons frame
        action_buttons_frame = ctk.CTkFrame(self.select_window, fg_color="transparent")
        action_buttons_frame.grid(row=3, column=0, padx=20, pady=(0, 20), sticky="ew")
        action_buttons_frame.grid_columnconfigure(0, weight=1)
        action_buttons_frame.grid_columnconfigure(1, weight=1)
        action_buttons_frame.grid_columnconfigure(2, weight=1)
        
        # Cancel button
        cancel_btn = ctk.CTkButton(
            action_buttons_frame,
            text="Cancel",
            command=self.select_window.destroy,
            fg_color="#999999",
            hover_color="#CCCCCC"
        )
        cancel_btn.grid(row=0, column=0, padx=5, pady=10)
        
        # Remove unselected columns button
        remove_cols_btn = ctk.CTkButton(
            action_buttons_frame,
            text="Remove Unselected Columns",
            command=self.remove_unselected_columns,
            fg_color="#FF6600",
            hover_color="#FF9933"
        )
        remove_cols_btn.grid(row=0, column=1, padx=5, pady=10)
        
        # Unpivot button
        unpivot_btn = ctk.CTkButton(
            action_buttons_frame,
            text="Perform Unpivot",
            command=self.perform_unpivot,
            fg_color=DARK_BLUE,
            hover_color=ACCENT_BLUE
        )
        unpivot_btn.grid(row=0, column=2, padx=5, pady=10)
    
    def select_all_columns(self):
        """Select all columns in the selection window"""
        for var in self.column_vars.values():
            var.set(True)
        self.update_status("All columns selected.")
    
    def deselect_all_columns(self):
        """Deselect all columns in the selection window"""
        for var in self.column_vars.values():
            var.set(False)
        self.update_status("All columns deselected.")
    
    def remove_unselected_columns(self):
        """Remove unselected columns, keeping only the selected ones"""
        try:
            # Get selected columns
            selected_cols = [col for col, var in self.column_vars.items() if var.get()]
            
            if not selected_cols:
                messagebox.showwarning("No Selection", "Please select at least one column to keep.")
                return
            
            # Save current state before removing columns
            self.save_state()
            
            # Keep only selected columns
            original_cols = list(self.df.columns)
            self.df = self.df[selected_cols]
            
            # Update the selection window checkboxes
            self.select_window.destroy()
            
            # Show confirmation
            removed_cols = [col for col in original_cols if col not in selected_cols]
            removed_count = len(removed_cols)
            
            # Update UI
            self.update_preview()
            self.update_status(f"Removed {removed_count} unselected columns. Kept {len(selected_cols)} columns.")
            
            # Show what was removed
            if removed_cols:
                removed_list = ", ".join(removed_cols[:5])  # Show first 5
                if len(removed_cols) > 5:
                    removed_list += f" and {len(removed_cols) - 5} more..."
                
                messagebox.showinfo("Columns Removed", 
                                  f"Removed {removed_count} column(s):\n\n{removed_list}\n\n"
                                  f"Kept {len(selected_cols)} column(s) for unpivoting.")
            
            # Re-open the column selection window with updated columns
            self.open_column_selection_window()
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to remove columns: {str(e)}")
            self.update_status("Error removing columns")
    
    def perform_unpivot(self):
        """Perform unpivoting based on selected columns"""
        try:
            # Get selected columns
            id_vars = [col for col, var in self.column_vars.items() if var.get()]
            
            if not id_vars:
                messagebox.showwarning("No Selection", "Please select at least one column to keep as ID.")
                return
            
            # Save current state before unpivoting
            self.save_state()
            
            # Get value columns (all other columns)
            value_vars = [col for col in self.df.columns if col not in id_vars]
            
            if not value_vars:
                messagebox.showwarning("No Value Columns", "All columns are selected as ID columns. Need at least one value column to unpivot.")
                return
            
            # Perform unpivot using melt
            self.df = self.df.melt(
                id_vars=id_vars,
                value_vars=value_vars,
                var_name="Variable",
                value_name="Value"
            )
            
            # Save state after unpivoting
            self.save_state()
            
            # Close selection window
            self.select_window.destroy()
            
            # Update UI
            self.update_preview()
            self.update_status(f"Unpivoting completed. New shape: {self.df.shape}. Undo available.")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to perform unpivot: {str(e)}")
            self.update_status("Error during unpivoting")
    
    def finish_current_file(self):
        """Finish processing the current file and move to next"""
        if self.df is not None:
            # Store the processed dataframe
            self.processed_dataframes.append({
                'filename': os.path.basename(self.files_to_process[self.current_file_index]),
                'dataframe': self.df.copy()
            })
            
            # Move to next file
            self.current_file_index += 1
            
            # Reset current file state
            self.df = None
            self.df_history = []
            self.current_step = -1
            
            # Update UI
            self.update_file_list_display()
            self.update_progress_label()
            self.undo_btn.configure(state="disabled")
            self.remove_current_btn.configure(state="disabled")
            
            if self.current_file_index < len(self.files_to_process):
                # Still have files to process
                next_filename = os.path.basename(self.files_to_process[self.current_file_index])
                self.current_file_label.configure(text=f"Ready to process next file: {next_filename}")
                self.preview_text.configure(state="normal")
                self.preview_text.delete("0.0", "end")
                self.preview_text.insert("0.0", f"File {self.current_file_index + 1} of {len(self.files_to_process)} ready to process.\nClick 'Process Next File' to continue.")
                self.preview_text.configure(state="disabled")
                
                # Update button states
                self.unpivot_btn.configure(state="disabled")
                self.clean_data_btn.configure(state="disabled")
                self.flip_layout_btn.configure(state="disabled")
                self.finish_file_btn.configure(state="disabled")
                
                self.update_status(f"File saved. {len(self.files_to_process) - self.current_file_index} file(s) remaining.")
            else:
                # All files processed
                self.handle_all_files_completed()
    
    def save_all_results(self):
        """Save all processed results as a single file with all data in one worksheet"""
        if not self.processed_dataframes:
            messagebox.showwarning("No Data", "No processed data to save.")
            return
        
        filetypes = [
            ("Excel files", "*.xlsx"),
            ("CSV files", "*.csv"),
        ]
        
        default_name = f"combined_unpivoted_results_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        
        file_path = filedialog.asksaveasfilename(
            title="Save All Results",
            defaultextension=".xlsx",
            initialfile=default_name,
            filetypes=filetypes
        )
        
        if file_path:
            try:
                self.update_status("Saving combined results...")
                
                # Combine all dataframes into one
                combined_df = pd.DataFrame()
                
                for item in self.processed_dataframes:
                    df = item['dataframe'].copy()
                    # Add a column to track the source file
                    df['Source_File'] = item['filename']
                    combined_df = pd.concat([combined_df, df], ignore_index=True)
                
                if file_path.endswith('.csv'):
                    # Save as CSV
                    combined_df.to_csv(file_path, index=False, encoding='utf-8-sig')
                else:
                    # Save as Excel with all data in one worksheet
                    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                        # Save combined data to main worksheet
                        combined_df.to_excel(writer, sheet_name='All_Results', index=False)
                        
                        # OPTIONAL: Also save a summary sheet
                        summary_data = []
                        for item in self.processed_dataframes:
                            df = item['dataframe']
                            summary_data.append({
                                'File Name': item['filename'],
                                'Rows': df.shape[0],
                                'Columns': df.shape[1]
                            })
                        
                        summary_df = pd.DataFrame(summary_data)
                        summary_df.to_excel(writer, sheet_name='Summary', index=False)
                
                # Calculate statistics
                total_rows = combined_df.shape[0]
                total_files = len(self.processed_dataframes)
                
                self.update_status(f"Combined results saved successfully: {os.path.basename(file_path)}")
                messagebox.showinfo("Success", 
                                  f"All results saved successfully!\n\n"
                                  f"• Total files combined: {total_files}\n"
                                  f"• Total rows in combined file: {total_rows}\n"
                                  f"• Location: {file_path}\n\n"
                                  f"All data has been combined into a single worksheet.")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save combined results: {str(e)}")
                self.update_status("Error saving combined results")

if __name__ == "__main__":
    app = MultiFileUnpivotApp()
    app.mainloop()

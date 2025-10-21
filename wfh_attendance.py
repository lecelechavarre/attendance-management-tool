import tkinter as tk
from tkinter import ttk, messagebox
import json
import os
from datetime import datetime
import pandas as pd
from typing import Dict, List, Optional
import subprocess
import sys

class WFHAttendanceApp:
    def __init__(self, root):
        self.root = root
        self.root.title("WFH Attendance")
        self.root.geometry("1000x700")
        self.root.configure(bg='#f0f2f5')
        
        # Center the window
        self.center_window()
        
        # Initialize data storage
        self.data_file = "attendance_data.json"
        self.sessions_file = "active_sessions.json"
        self.attendance_data = self.load_data()
        self.active_sessions = self.load_sessions()
        
        # Current user session
        self.current_user_id = None
        self.current_session_id = None
        
        # Configure styles
        self.setup_styles()
        
        # Create UI
        self.create_main_frame()
        self.create_login_section()
        self.create_attendance_section()
        self.create_export_section()
        self.create_records_section()
        self.create_sessions_section()
        
        # Update records display
        self.update_records_display()
        self.update_sessions_display()
    
    def center_window(self):
        """Center the window on the screen"""
        self.root.update_idletasks()
        width = 1000
        height = 700
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')
    
    def setup_styles(self):
        """Configure modern styles"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Configure custom styles
        style.configure('Title.TLabel', 
                       font=('Arial', 16, 'bold'),
                       background='#f0f2f5',
                       foreground='#2c3e50')
        
        style.configure('Card.TFrame',
                       background='white',
                       relief='raised',
                       borderwidth=1)
        
        style.configure('Primary.TButton',
                       font=('Arial', 10, 'bold'),
                       background='#3498db',
                       foreground='white')
        
        style.configure('Success.TButton',
                       font=('Arial', 10, 'bold'),
                       background='#27ae60',
                       foreground='white')
        
        style.configure('Warning.TButton',
                       font=('Arial', 10, 'bold'),
                       background='#e67e22',
                       foreground='white')
        
        style.configure('Danger.TButton',
                       font=('Arial', 10, 'bold'),
                       background='#e74c3c',
                       foreground='white')
    
    def create_main_frame(self):
        """Create main container frame"""
        self.main_frame = ttk.Frame(self.root, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Title
        title_label = ttk.Label(
            self.main_frame,
            text="WFH Attendance System",
            style='Title.TLabel'
        )
        title_label.pack(pady=(0, 20))
    
    def create_login_section(self):
        """Create user login section"""
        login_frame = ttk.LabelFrame(
            self.main_frame,
            text="User Login",
            padding="15"
        )
        login_frame.pack(fill=tk.X, pady=(0, 10))
        
        # User ID
        ttk.Label(login_frame, text="User ID:").grid(
            row=0, column=0, padx=(0, 10), pady=5, sticky=tk.W
        )
        self.user_id_var = tk.StringVar()
        self.user_id_entry = ttk.Entry(
            login_frame, 
            textvariable=self.user_id_var,
            width=20,
            font=('Arial', 10)
        )
        self.user_id_entry.grid(row=0, column=1, padx=(0, 20), pady=5)
        
        # User Name
        ttk.Label(login_frame, text="User Name:").grid(
            row=0, column=2, padx=(0, 10), pady=5, sticky=tk.W
        )
        self.user_name_var = tk.StringVar()
        self.user_name_entry = ttk.Entry(
            login_frame, 
            textvariable=self.user_name_var,
            width=25,
            font=('Arial', 10)
        )
        self.user_name_entry.grid(row=0, column=3, padx=(0, 20), pady=5)
        
        # Login Button
        self.login_btn = ttk.Button(
            login_frame,
            text="Login / Register",
            command=self.handle_login,
            style='Primary.TButton'
        )
        self.login_btn.grid(row=0, column=4, padx=(0, 10), pady=5)
        
        # Logout Button
        self.logout_btn = ttk.Button(
            login_frame,
            text="Logout",
            command=self.handle_logout,
            style='Danger.TButton',
            state=tk.DISABLED
        )
        self.logout_btn.grid(row=0, column=5, padx=(0, 10), pady=5)
        
        # Status Label
        self.login_status_var = tk.StringVar(value="Please login to continue")
        self.login_status_label = ttk.Label(
            login_frame,
            textvariable=self.login_status_var,
            foreground='#e74c3c',
            font=('Arial', 9)
        )
        self.login_status_label.grid(row=1, column=0, columnspan=6, sticky=tk.W, pady=(5, 0))
    
    def create_attendance_section(self):
        """Create attendance recording section"""
        self.attendance_frame = ttk.LabelFrame(
            self.main_frame,
            text="Attendance",
            padding="15"
        )
        self.attendance_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Time display
        self.time_display_var = tk.StringVar(value="Current Time: --:--:--")
        time_label = ttk.Label(
            self.attendance_frame,
            textvariable=self.time_display_var,
            font=('Arial', 12, 'bold'),
            foreground='#2c3e50'
        )
        time_label.pack(pady=(0, 10))
        
        # Buttons frame
        btn_frame = ttk.Frame(self.attendance_frame)
        btn_frame.pack(fill=tk.X)
        
        # Time In Button
        self.time_in_btn = ttk.Button(
            btn_frame,
            text="Time In",
            command=self.time_in,
            style='Success.TButton',
            state=tk.DISABLED
        )
        self.time_in_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Time Out Button
        self.time_out_btn = ttk.Button(
            btn_frame,
            text="Time Out",
            command=self.time_out,
            style='Warning.TButton',
            state=tk.DISABLED
        )
        self.time_out_btn.pack(side=tk.LEFT)
        
        # Status display
        self.attendance_status_var = tk.StringVar(value="No active session")
        self.attendance_status_label = ttk.Label(
            self.attendance_frame,
            textvariable=self.attendance_status_var,
            font=('Arial', 10),
            foreground='#7f8c8d'
        )
        self.attendance_status_label.pack(pady=(10, 0))
        
        # Start clock update
        self.update_clock()
    
    def create_export_section(self):
        """Create export controls section"""
        export_frame = ttk.Frame(self.main_frame)
        export_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Export Button
        self.export_btn = ttk.Button(
            export_frame,
            text="Export to Excel",
            command=self.export_to_excel,
            style='Primary.TButton'
        )
        self.export_btn.pack(side=tk.RIGHT)
        
        # Export Path Label
        self.export_path_var = tk.StringVar(value="Export path will appear here")
        self.export_path_label = ttk.Label(
            export_frame,
            textvariable=self.export_path_var,
            font=('Arial', 8),
            foreground='#7f8c8d'
        )
        self.export_path_label.pack(side=tk.RIGHT, padx=(0, 10))
    
    def create_records_section(self):
        """Create attendance records display"""
        # Create notebook for tabs
        self.notebook = ttk.Notebook(self.main_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Records tab
        records_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(records_frame, text="Attendance Records")
        
        # Create treeview with scrollbar
        tree_frame = ttk.Frame(records_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Treeview
        columns = ('user_id', 'user_name', 'date', 'time_in', 'time_out', 'duration')
        self.records_tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show='headings',
            yscrollcommand=scrollbar.set
        )
        
        # Configure columns
        self.records_tree.heading('user_id', text='User ID')
        self.records_tree.heading('user_name', text='User Name')
        self.records_tree.heading('date', text='Date')
        self.records_tree.heading('time_in', text='Time In')
        self.records_tree.heading('time_out', text='Time Out')
        self.records_tree.heading('duration', text='Duration')
        
        self.records_tree.column('user_id', width=80)
        self.records_tree.column('user_name', width=120)
        self.records_tree.column('date', width=100)
        self.records_tree.column('time_in', width=100)
        self.records_tree.column('time_out', width=100)
        self.records_tree.column('duration', width=80)
        
        self.records_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.records_tree.yview)
    
    def create_sessions_section(self):
        """Create active sessions display tab"""
        sessions_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(sessions_frame, text="Active Sessions")
        
        # Create treeview with scrollbar
        tree_frame = ttk.Frame(sessions_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(tree_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Treeview
        columns = ('session_id', 'user_id', 'user_name', 'date', 'time_in')
        self.sessions_tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show='headings',
            yscrollcommand=scrollbar.set
        )
        
        # Configure columns
        self.sessions_tree.heading('session_id', text='Session ID')
        self.sessions_tree.heading('user_id', text='User ID')
        self.sessions_tree.heading('user_name', text='User Name')
        self.sessions_tree.heading('date', text='Date')
        self.sessions_tree.heading('time_in', text='Time In')
        
        self.sessions_tree.column('session_id', width=100)
        self.sessions_tree.column('user_id', width=80)
        self.sessions_tree.column('user_name', width=120)
        self.sessions_tree.column('date', width=100)
        self.sessions_tree.column('time_in', width=100)
        
        self.sessions_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.sessions_tree.yview)
        
        # Force time out button (admin feature)
        force_out_btn = ttk.Button(
            sessions_frame,
            text="Force Time Out Selected Session",
            command=self.force_time_out,
            style='Danger.TButton'
        )
        force_out_btn.pack(pady=(10, 0))
    
    def update_clock(self):
        """Update the current time display"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.time_display_var.set(f"Current Time: {current_time}")
        self.root.after(1000, self.update_clock)
    
    def generate_session_id(self, user_id: str) -> str:
        """Generate unique session ID"""
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        return f"{user_id}_{timestamp}"
    
    def handle_login(self):
        """Handle user login/registration"""
        user_id = self.user_id_var.get().strip()
        user_name = self.user_name_var.get().strip()
        
        if not user_id or not user_name:
            messagebox.showerror("Error", "Please enter both User ID and User Name")
            return
        
        self.current_user_id = user_id
        
        # Update status
        self.login_status_var.set(f"Logged in as: {user_name} ({user_id})")
        self.login_status_label.configure(foreground='#27ae60')
        
        # Enable/disable buttons
        self.login_btn.config(state=tk.DISABLED)
        self.logout_btn.config(state=tk.NORMAL)
        self.time_in_btn.config(state=tk.NORMAL)
        self.user_id_entry.config(state=tk.DISABLED)
        self.user_name_entry.config(state=tk.DISABLED)
        
        # Check for active session
        self.check_active_session()
    
    def handle_logout(self):
        """Handle user logout"""
        self.current_user_id = None
        self.current_session_id = None
        
        # Update status
        self.login_status_var.set("Please login to continue")
        self.login_status_label.configure(foreground='#e74c3c')
        self.attendance_status_var.set("No active session")
        
        # Enable/disable buttons
        self.login_btn.config(state=tk.NORMAL)
        self.logout_btn.config(state=tk.DISABLED)
        self.time_in_btn.config(state=tk.DISABLED)
        self.time_out_btn.config(state=tk.DISABLED)
        self.user_id_entry.config(state=tk.NORMAL)
        self.user_name_entry.config(state=tk.NORMAL)
    
    def check_active_session(self):
        """Check if user has an active time-in session"""
        user_sessions = [s for s in self.active_sessions if s['user_id'] == self.current_user_id]
        
        if user_sessions:
            session = user_sessions[0]
            self.current_session_id = session['session_id']
            self.attendance_status_var.set(
                f"Active session: Time In at {session['time_in']}"
            )
            self.time_in_btn.config(state=tk.DISABLED)
            self.time_out_btn.config(state=tk.NORMAL)
        else:
            self.current_session_id = None
            self.attendance_status_var.set("Ready for Time In")
            self.time_in_btn.config(state=tk.NORMAL)
            self.time_out_btn.config(state=tk.DISABLED)
    
    def time_in(self):
        """Record time in with session management"""
        if not self.current_user_id:
            messagebox.showerror("Error", "Please login first")
            return
        
        # Check if user already has active session
        user_sessions = [s for s in self.active_sessions if s['user_id'] == self.current_user_id]
        if user_sessions:
            messagebox.showerror("Error", "You already have an active session!")
            return
        
        current_time = datetime.now().strftime("%H:%M:%S")
        today = datetime.now().strftime("%Y-%m-%d")
        
        # Generate session ID
        session_id = self.generate_session_id(self.current_user_id)
        
        # Create session record
        session_record = {
            'session_id': session_id,
            'user_id': self.current_user_id,
            'user_name': self.user_name_var.get(),
            'date': today,
            'time_in': current_time
        }
        
        # Add to active sessions
        self.active_sessions.append(session_record)
        self.current_session_id = session_id
        
        # Save sessions
        self.save_sessions()
        
        self.attendance_status_var.set(f"Time In recorded at {current_time}")
        self.time_in_btn.config(state=tk.DISABLED)
        self.time_out_btn.config(state=tk.NORMAL)
        
        self.update_sessions_display()
        messagebox.showinfo("Success", "Time In recorded successfully!")
    
    def time_out(self):
        """Record time out with session management"""
        if not self.current_session_id:
            messagebox.showerror("Error", "No active session found")
            return
        
        # Find the session
        session_index = None
        session_data = None
        for i, session in enumerate(self.active_sessions):
            if session['session_id'] == self.current_session_id:
                session_index = i
                session_data = session
                break
        
        if session_index is None:
            messagebox.showerror("Error", "Session not found")
            return
        
        current_time = datetime.now().strftime("%H:%M:%S")
        
        # Create attendance record
        record = {
            'user_id': session_data['user_id'],
            'user_name': session_data['user_name'],
            'date': session_data['date'],
            'time_in': session_data['time_in'],
            'time_out': current_time,
            'duration': self.calculate_duration(session_data['time_in'], current_time)
        }
        
        # Add to attendance data
        self.attendance_data.append(record)
        
        # Remove from active sessions
        self.active_sessions.pop(session_index)
        self.current_session_id = None
        
        # Save data
        self.save_data()
        self.save_sessions()
        
        self.attendance_status_var.set(f"Time Out recorded at {current_time}")
        self.time_in_btn.config(state=tk.NORMAL)
        self.time_out_btn.config(state=tk.DISABLED)
        
        self.update_records_display()
        self.update_sessions_display()
        messagebox.showinfo("Success", "Time Out recorded successfully!")
    
    def force_time_out(self):
        """Force time out for selected session (admin feature)"""
        selected = self.sessions_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select a session to force time out")
            return
        
        item = selected[0]
        session_id = self.sessions_tree.item(item, 'values')[0]
        
        # Find the session
        session_index = None
        session_data = None
        for i, session in enumerate(self.active_sessions):
            if session['session_id'] == session_id:
                session_index = i
                session_data = session
                break
        
        if session_index is None:
            messagebox.showerror("Error", "Session not found")
            return
        
        current_time = datetime.now().strftime("%H:%M:%S")
        
        # Create attendance record
        record = {
            'user_id': session_data['user_id'],
            'user_name': session_data['user_name'],
            'date': session_data['date'],
            'time_in': session_data['time_in'],
            'time_out': current_time,
            'duration': self.calculate_duration(session_data['time_in'], current_time)
        }
        
        # Add to attendance data
        self.attendance_data.append(record)
        
        # Remove from active sessions
        self.active_sessions.pop(session_index)
        
        # If it's the current user's session, update their status
        if self.current_session_id == session_id:
            self.current_session_id = None
            self.attendance_status_var.set("Ready for Time In")
            self.time_in_btn.config(state=tk.NORMAL)
            self.time_out_btn.config(state=tk.DISABLED)
        
        # Save data
        self.save_data()
        self.save_sessions()
        
        self.update_records_display()
        self.update_sessions_display()
        messagebox.showinfo("Success", "Session force timed out successfully!")
    
    def calculate_duration(self, time_in: str, time_out: str) -> str:
        """Calculate duration between time in and time out"""
        time_in_dt = datetime.strptime(time_in, "%H:%M:%S")
        time_out_dt = datetime.strptime(time_out, "%H:%M:%S")
        duration = time_out_dt - time_in_dt
        
        hours, remainder = divmod(duration.total_seconds(), 3600)
        minutes, seconds = divmod(remainder, 60)
        return f"{int(hours):02d}:{int(minutes):02d}"
    
    def export_to_excel(self):
        """Export attendance data to Excel and show file path"""
        if not self.attendance_data:
            messagebox.showwarning("Warning", "No attendance data to export")
            return
        
        try:
            # Create DataFrame
            df = pd.DataFrame(self.attendance_data)
            
            # Generate filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"wfh_attendance_{timestamp}.xlsx"
            
            # Get absolute path
            abs_path = os.path.abspath(filename)
            
            # Export to Excel
            df.to_excel(filename, index=False, engine='openpyxl')
            
            # Update export path label
            self.export_path_var.set(f"Exported to: {abs_path}")
            
            messagebox.showinfo("Success", f"Data exported to:\n{abs_path}")
            
            # Ask if user wants to open the file
            if messagebox.askyesno("Open File", "Do you want to open the Excel file?"):
                self.open_file(abs_path)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data: {str(e)}")
    
    def open_file(self, filepath: str):
        """Open file with default application"""
        try:
            if sys.platform == "win32":
                os.startfile(filepath)
            elif sys.platform == "darwin":  # macOS
                subprocess.run(["open", filepath])
            else:  # linux
                subprocess.run(["xdg-open", filepath])
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open file: {str(e)}")
    
    def update_records_display(self):
        """Update the records treeview"""
        # Clear existing records
        for item in self.records_tree.get_children():
            self.records_tree.delete(item)
        
        # Add records (show latest first)
        for record in reversed(self.attendance_data):
            self.records_tree.insert(
                '', tk.END,
                values=(
                    record['user_id'],
                    record['user_name'],
                    record['date'],
                    record['time_in'],
                    record['time_out'],
                    record['duration']
                )
            )
    
    def update_sessions_display(self):
        """Update the active sessions treeview"""
        # Clear existing sessions
        for item in self.sessions_tree.get_children():
            self.sessions_tree.delete(item)
        
        # Add active sessions
        for session in self.active_sessions:
            self.sessions_tree.insert(
                '', tk.END,
                values=(
                    session['session_id'],
                    session['user_id'],
                    session['user_name'],
                    session['date'],
                    session['time_in']
                )
            )
    
    def load_data(self) -> List[Dict]:
        """Load attendance data from JSON file"""
        try:
            if os.path.exists(self.data_file):
                with open(self.data_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading data: {e}")
        return []
    
    def load_sessions(self) -> List[Dict]:
        """Load active sessions from JSON file"""
        try:
            if os.path.exists(self.sessions_file):
                with open(self.sessions_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading sessions: {e}")
        return []
    
    def save_data(self):
        """Save attendance data to JSON file"""
        try:
            with open(self.data_file, 'w') as f:
                json.dump(self.attendance_data, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save data: {str(e)}")
    
    def save_sessions(self):
        """Save active sessions to JSON file"""
        try:
            with open(self.sessions_file, 'w') as f:
                json.dump(self.active_sessions, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save sessions: {str(e)}")

def main():
    root = tk.Tk()
    app = WFHAttendanceApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
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
        self.export_history_file = "export_history.json"
        self.users_file = "registered_users.json"
        self.archive_file = "deleted_users_archive.json"
        self.admin_file = "admin_users.json"
        self.roles_file = "roles_users.json"
        self.attendance_data = self.load_data()
        self.active_sessions = self.load_sessions()
        self.export_history = self.load_export_history()
        self.registered_users = self.load_registered_users()
        self.deleted_users_archive = self.load_archive()
        self.admin_users = self.load_admin_users()
        self.roles_users = self.load_roles_users()
        
        # Current user session
        self.current_user_id = None
        self.current_session_id = None
        self.user_role = None  # 'admin', 'roles', 'regular'
        
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
        
        # Initially hide features based on role
        self.toggle_features_based_on_role()
    
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
        
        style.configure('Info.TButton',
                       font=('Arial', 10, 'bold'),
                       background='#9b59b6',
                       foreground='white')
        
        style.configure('Admin.TButton',
                       font=('Arial', 10, 'bold'),
                       background='#f39c12',
                       foreground='white')
        
        style.configure('Roles.TButton',
                       font=('Arial', 10, 'bold'),
                       background='#8e44ad',
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
            text="Login",
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
        
        # Manage Users Button (Admin only)
        self.manage_users_btn = ttk.Button(
            login_frame,
            text="Manage Users",
            command=self.manage_users,
            style='Info.TButton',
            state=tk.DISABLED
        )
        self.manage_users_btn.grid(row=0, column=6, padx=(0, 10), pady=5)
        
        # Status Label
        self.login_status_var = tk.StringVar(value="Please login to continue")
        self.login_status_label = ttk.Label(
            login_frame,
            textvariable=self.login_status_var,
            foreground='#e74c3c',
            font=('Arial', 9)
        )
        self.login_status_label.grid(row=1, column=0, columnspan=7, sticky=tk.W, pady=(5, 0))
    
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
        
        # Auto Time In Button (Roles User only)
        self.auto_time_in_btn = ttk.Button(
            btn_frame,
            text="Auto New Session",
            command=self.auto_new_session,
            style='Info.TButton',
            state=tk.DISABLED
        )
        self.auto_time_in_btn.pack(side=tk.LEFT, padx=(10, 0))
        
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
        self.export_frame = ttk.Frame(self.main_frame)
        
        # Export Button (Admin and Roles only)
        self.export_btn = ttk.Button(
            self.export_frame,
            text="Export to Excel & Start New Session",
            command=self.export_to_excel,
            style='Roles.TButton'
        )
        self.export_btn.pack(side=tk.RIGHT)
        
        # Export Path Label
        self.export_path_var = tk.StringVar(value="Export path will appear here")
        self.export_path_label = ttk.Label(
            self.export_frame,
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
        self.notebook.add(records_frame, text="My Attendance Records")
        
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
        
        # Force time out button (Admin only)
        self.force_out_btn = ttk.Button(
            sessions_frame,
            text="Force Time Out Selected Session",
            command=self.force_time_out,
            style='Danger.TButton'
        )
    
    def update_clock(self):
        """Update the current time display"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.time_display_var.set(f"Current Time: {current_time}")
        self.root.after(1000, self.update_clock)
    
    def toggle_features_based_on_role(self):
        """Show/hide features based on user role"""
        if self.user_role == 'admin':
            # Admin: Show user management, export, hide attendance
            self.export_frame.pack(fill=tk.X, pady=(0, 10))
            self.manage_users_btn.config(state=tk.NORMAL)
            self.force_out_btn.pack(pady=(10, 0))
            self.attendance_frame.pack_forget()  # Hide attendance for admin
            self.auto_time_in_btn.pack_forget()
        elif self.user_role == 'roles':
            # Roles User: Show export and attendance, hide user management
            self.export_frame.pack(fill=tk.X, pady=(0, 10))
            self.manage_users_btn.config(state=tk.DISABLED)
            self.force_out_btn.pack_forget()
            self.attendance_frame.pack(fill=tk.X, pady=(0, 10))
            self.auto_time_in_btn.pack(side=tk.LEFT, padx=(10, 0))
        elif self.user_role == 'regular':
            # Regular User: Show only attendance, hide everything else
            self.export_frame.pack_forget()
            self.manage_users_btn.config(state=tk.DISABLED)
            self.force_out_btn.pack_forget()
            self.attendance_frame.pack(fill=tk.X, pady=(0, 10))
            self.auto_time_in_btn.pack_forget()
        else:
            # Not logged in: Hide everything
            self.export_frame.pack_forget()
            self.manage_users_btn.config(state=tk.DISABLED)
            self.force_out_btn.pack_forget()
            self.attendance_frame.pack_forget()
    
    def get_user_role(self, user_id: str) -> str:
        """Get user role: admin, roles, or regular"""
        for admin in self.admin_users:
            if admin['user_id'].lower() == user_id.lower():
                return 'admin'
        for roles_user in self.roles_users:
            if roles_user['user_id'].lower() == user_id.lower():
                return 'roles'
        return 'regular'
    
    def generate_session_id(self, user_id: str) -> str:
        """Generate unique session ID"""
        timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
        return f"{user_id}_{timestamp}"
    
    def check_duplicate_user(self, user_id: str, user_name: str) -> tuple:
        """Check if User ID or User Name already exists"""
        for user in self.registered_users:
            if user['user_id'].lower() == user_id.lower():
                return True, f"User ID '{user_id}' is already registered to '{user['user_name']}'"
            if user['user_name'].lower() == user_name.lower():
                return True, f"User Name '{user_name}' is already registered to User ID '{user['user_id']}'"
        return False, ""
    
    def register_new_user(self, user_id: str, user_name: str, role: str = 'regular'):
        """Register a new user with specific role"""
        new_user = {
            'user_id': user_id,
            'user_name': user_name,
            'registered_date': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'role': role
        }
        self.registered_users.append(new_user)
        self.save_registered_users()
        
        # Add to specific role list
        if role == 'admin':
            admin_user = {
                'user_id': user_id,
                'user_name': user_name,
                'admin_since': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            self.admin_users.append(admin_user)
            self.save_admin_users()
        elif role == 'roles':
            roles_user = {
                'user_id': user_id,
                'user_name': user_name,
                'roles_since': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            }
            self.roles_users.append(roles_user)
            self.save_roles_users()
    
    def handle_login(self):
        """Handle user login"""
        user_id = self.user_id_var.get().strip()
        user_name = self.user_name_var.get().strip()
        
        if not user_id or not user_name:
            messagebox.showerror("Error", "Please enter both User ID and User Name")
            return
        
        # Check if user exists and credentials match
        user_exists = False
        correct_credentials = False
        user_data = None
        
        for user in self.registered_users:
            if user['user_id'].lower() == user_id.lower():
                user_exists = True
                if user['user_name'].lower() == user_name.lower():
                    correct_credentials = True
                    user_data = user
                    break
                else:
                    messagebox.showerror(
                        "Login Error", 
                        f"User ID '{user_id}' is registered to '{user['user_name']}'. "
                        f"Please use the correct User Name."
                    )
                    return
        
        if not user_exists:
            messagebox.showerror("Login Error", "User ID not found. Please contact administrator for registration.")
            return
        
        if correct_credentials and user_data:
            # Determine user role
            self.user_role = self.get_user_role(user_id)
            role_msg = "Admin" if self.user_role == 'admin' else "Roles User" if self.user_role == 'roles' else "Regular User"
            
            self.current_user_id = user_id
            
            # Update status with role indicator
            self.login_status_var.set(f"Logged in as: {user_name} ({user_id}) - {role_msg}")
            self.login_status_label.configure(foreground='#27ae60')
            
            # Enable/disable buttons based on role
            self.login_btn.config(state=tk.DISABLED)
            self.logout_btn.config(state=tk.NORMAL)
            self.user_id_entry.config(state=tk.DISABLED)
            self.user_name_entry.config(state=tk.DISABLED)
            
            # Toggle features based on role
            self.toggle_features_based_on_role()
            
            # For regular and roles users, enable time in/out and check session
            if self.user_role in ['regular', 'roles']:
                self.time_in_btn.config(state=tk.NORMAL)
                self.auto_time_in_btn.config(state=tk.NORMAL)
                self.check_active_session()
            
            # Update records display to show only current user's records for non-admin users
            if self.user_role != 'admin':
                self.update_records_display()
            
            messagebox.showinfo("Login Successful", f"Welcome {user_name}! ({role_msg})")
    
    def manage_users(self):
        """Show user management window (Admin only)"""
        if self.user_role != 'admin':
            messagebox.showerror("Access Denied", "Only administrators can manage users.")
            return
            
        # Create new window
        users_window = tk.Toplevel(self.root)
        users_window.title("User Management - Admin")
        users_window.geometry("800x600")
        users_window.configure(bg='#f0f2f5')
        
        # Center the window
        users_window.update_idletasks()
        width = 800
        height = 600
        x = (users_window.winfo_screenwidth() // 2) - (width // 2)
        y = (users_window.winfo_screenheight() // 2) - (height // 2)
        users_window.geometry(f'{width}x{height}+{x}+{y}')
        
        # Title
        title_label = ttk.Label(
            users_window,
            text="User Management - Admin Panel",
            font=('Arial', 14, 'bold'),
            background='#f0f2f5',
            foreground='#2c3e50'
        )
        title_label.pack(pady=10)
        
        # Registration Frame
        reg_frame = ttk.LabelFrame(users_window, text="Register New User", padding="10")
        reg_frame.pack(fill=tk.X, padx=20, pady=10)
        
        # User ID
        ttk.Label(reg_frame, text="User ID:").grid(row=0, column=0, padx=(0, 10), pady=5, sticky=tk.W)
        new_user_id_var = tk.StringVar()
        new_user_id_entry = ttk.Entry(reg_frame, textvariable=new_user_id_var, width=15)
        new_user_id_entry.grid(row=0, column=1, padx=(0, 20), pady=5)
        
        # User Name
        ttk.Label(reg_frame, text="User Name:").grid(row=0, column=2, padx=(0, 10), pady=5, sticky=tk.W)
        new_user_name_var = tk.StringVar()
        new_user_name_entry = ttk.Entry(reg_frame, textvariable=new_user_name_var, width=20)
        new_user_name_entry.grid(row=0, column=3, padx=(0, 20), pady=5)
        
        # Role Selection
        ttk.Label(reg_frame, text="Role:").grid(row=0, column=4, padx=(0, 10), pady=5, sticky=tk.W)
        role_var = tk.StringVar(value="regular")
        role_combo = ttk.Combobox(reg_frame, textvariable=role_var, values=["regular", "roles", "admin"], state="readonly", width=10)
        role_combo.grid(row=0, column=5, padx=(0, 20), pady=5)
        
        def register_new_user_admin():
            user_id = new_user_id_var.get().strip()
            user_name = new_user_name_var.get().strip()
            role = role_var.get()
            
            if not user_id or not user_name:
                messagebox.showerror("Error", "Please enter both User ID and User Name")
                return
            
            # Check for duplicates
            is_duplicate, error_message = self.check_duplicate_user(user_id, user_name)
            if is_duplicate:
                messagebox.showerror("Duplicate User", error_message)
                return
            
            # Register new user
            self.register_new_user(user_id, user_name, role)
            messagebox.showinfo("Success", f"User '{user_name}' ({user_id}) registered as {role} successfully!")
            
            # Clear fields
            new_user_id_var.set("")
            new_user_name_var.set("")
            role_var.set("regular")
            
            # Refresh user list
            refresh_user_list()
        
        register_btn = ttk.Button(reg_frame, text="Register User", command=register_new_user_admin, style='Success.TButton')
        register_btn.grid(row=0, column=6, padx=(0, 10), pady=5)
        
        # Users List Frame
        list_frame = ttk.LabelFrame(users_window, text="Registered Users", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=10)
        
        # Create treeview with scrollbar
        tree_frame = ttk.Frame(list_frame)
        tree_frame.pack(fill=tk.BOTH, expand=True)
        
        scrollbar = ttk.Scrollbar(tree_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Treeview
        columns = ('user_id', 'user_name', 'role', 'registered_date')
        users_tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show='headings',
            yscrollcommand=scrollbar.set
        )
        
        # Configure columns
        users_tree.heading('user_id', text='User ID')
        users_tree.heading('user_name', text='User Name')
        users_tree.heading('role', text='Role')
        users_tree.heading('registered_date', text='Registration Date')
        
        users_tree.column('user_id', width=100)
        users_tree.column('user_name', width=150)
        users_tree.column('role', width=80)
        users_tree.column('registered_date', width=150)
        
        users_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=users_tree.yview)
        
        def refresh_user_list():
            # Clear existing items
            for item in users_tree.get_children():
                users_tree.delete(item)
            
            # Add users to treeview
            for user in self.registered_users:
                users_tree.insert('', tk.END, values=(
                    user['user_id'],
                    user['user_name'],
                    user.get('role', 'regular'),
                    user['registered_date']
                ))
        
        def delete_selected_user():
            selected = users_tree.selection()
            if not selected:
                messagebox.showwarning("Warning", "Please select a user to delete")
                return
            
            item = selected[0]
            user_id = users_tree.item(item, 'values')[0]
            user_name = users_tree.item(item, 'values')[1]
            
            # Prevent self-deletion
            if user_id == self.current_user_id:
                messagebox.showerror("Error", "You cannot delete your own account")
                return
            
            # Confirm deletion
            confirm = messagebox.askyesno("Confirm Deletion", f"Are you sure you want to delete user '{user_name}' ({user_id})?")
            if confirm:
                # Remove from registered users
                self.registered_users = [u for u in self.registered_users if u['user_id'] != user_id]
                self.save_registered_users()
                
                # Remove from role-specific lists
                self.admin_users = [a for a in self.admin_users if a['user_id'] != user_id]
                self.save_admin_users()
                
                self.roles_users = [r for r in self.roles_users if r['user_id'] != user_id]
                self.save_roles_users()
                
                # Refresh list
                refresh_user_list()
                messagebox.showinfo("Success", f"User '{user_name}' deleted successfully")
        
        # Buttons frame
        btn_frame = ttk.Frame(list_frame)
        btn_frame.pack(fill=tk.X, pady=10)
        
        delete_btn = ttk.Button(btn_frame, text="Delete Selected User", command=delete_selected_user, style='Danger.TButton')
        delete_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        refresh_btn = ttk.Button(btn_frame, text="Refresh List", command=refresh_user_list, style='Primary.TButton')
        refresh_btn.pack(side=tk.LEFT)
        
        close_btn = ttk.Button(users_window, text="Close", command=users_window.destroy, style='Primary.TButton')
        close_btn.pack(pady=10)
        
        # Initial load
        refresh_user_list()
    
    def handle_logout(self):
        """Handle user logout"""
        # Confirm logout if user has active session
        if self.current_session_id:
            confirm = messagebox.askyesno(
                "Logout with Active Session",
                "You have an active Time In session. Are you sure you want to logout?\n\n"
                "You can login again later to complete your Time Out."
            )
            if not confirm:
                return
        
        self.current_user_id = None
        self.current_session_id = None
        self.user_role = None
        
        # Update status
        self.login_status_var.set("Please login to continue")
        self.login_status_label.configure(foreground='#e74c3c')
        self.attendance_status_var.set("No active session")
        
        # Enable/disable buttons
        self.login_btn.config(state=tk.NORMAL)
        self.logout_btn.config(state=tk.DISABLED)
        self.time_in_btn.config(state=tk.DISABLED)
        self.time_out_btn.config(state=tk.DISABLED)
        self.auto_time_in_btn.config(state=tk.DISABLED)
        self.user_id_entry.config(state=tk.NORMAL)
        self.user_name_entry.config(state=tk.NORMAL)
        
        # Hide all features
        self.toggle_features_based_on_role()
        
        # Clear input fields
        self.user_id_var.set("")
        self.user_name_var.set("")
        
        messagebox.showinfo("Logout Successful", "You have been logged out successfully!")
    
    def check_active_session(self):
        """Check if user has an active time-in session"""
        if not self.current_user_id:
            return
            
        user_sessions = [s for s in self.active_sessions if s['user_id'] == self.current_user_id]
        
        if user_sessions:
            session = user_sessions[0]
            self.current_session_id = session['session_id']
            self.attendance_status_var.set(f"Active session: Time In at {session['time_in']}")
            self.time_in_btn.config(state=tk.DISABLED)
            self.time_out_btn.config(state=tk.NORMAL)
            self.auto_time_in_btn.config(state=tk.DISABLED)
        else:
            self.current_session_id = None
            self.attendance_status_var.set("Ready for Time In")
            self.time_in_btn.config(state=tk.NORMAL)
            self.time_out_btn.config(state=tk.DISABLED)
            self.auto_time_in_btn.config(state=tk.NORMAL)
    
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
        
        self.create_new_session()
    
    def create_new_session(self):
        """Create a new time-in session for the current user"""
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
        # Update button states
        self.time_in_btn.config(state=tk.DISABLED)
        self.time_out_btn.config(state=tk.NORMAL)
        self.auto_time_in_btn.config(state=tk.DISABLED)
        
        self.update_sessions_display()
        messagebox.showinfo("Success", "Time In recorded successfully!")
    
    def auto_new_session(self):
        """Automatically create a new session without time out (Roles User only)"""
        if self.user_role != 'roles':
            messagebox.showerror("Access Denied", "Only Roles Users can create auto sessions.")
            return
            
        if not self.current_user_id:
            messagebox.showerror("Error", "Please login first")
            return
        
        # Check if user already has active session
        user_sessions = [s for s in self.active_sessions if s['user_id'] == self.current_user_id]
        if user_sessions:
            messagebox.showerror("Error", "You already have an active session!")
            return
        
        # Confirm with user
        confirm = messagebox.askyesno(
            "Auto New Session", 
            "This will create a new Time In session without Time Out.\n\n"
            "Are you sure you want to start a new session?"
        )
        
        if confirm:
            self.create_new_session()
    
    def time_out(self):
        """Record time out with user validation"""
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
        
        # Validate user credentials against the session
        current_user_id = self.user_id_var.get().strip()
        current_user_name = self.user_name_var.get().strip()
        
        session_user_id = session_data['user_id']
        session_user_name = session_data['user_name']
        
        # Check if credentials match the session
        if current_user_id != session_user_id or current_user_name != session_user_name:
            self.show_validation_error(session_user_id, session_user_name)
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
        # Update button states
        self.time_in_btn.config(state=tk.NORMAL)
        self.time_out_btn.config(state=tk.DISABLED)
        self.auto_time_in_btn.config(state=tk.NORMAL)
        
        self.update_records_display()
        self.update_sessions_display()
        messagebox.showinfo("Success", "Time Out recorded successfully!")
    
    def show_validation_error(self, session_user_id: str, session_user_name: str):
        """Show error message when user credentials don't match the session"""
        error_message = (
            "User credentials do not match the active session!\n\n"
            f"Active session belongs to:\n"
            f"User ID: {session_user_id}\n"
            f"User Name: {session_user_name}\n\n"
            "Please enter the correct User ID and User Name that match the active session."
        )
        
        messagebox.showerror("Validation Error", error_message)
    
    def force_time_out(self):
        """Force time out for selected session (Admin only)"""
        if self.user_role != 'admin':
            messagebox.showerror("Access Denied", "Only administrators can force time out sessions.")
            return
            
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
        """Export attendance data to Excel (Admin and Roles only)"""
        if self.user_role not in ['admin', 'roles']:
            messagebox.showerror("Access Denied", "Only Admin and Roles Users can export data to Excel.")
            return
            
        if not self.attendance_data:
            messagebox.showwarning("Warning", "No attendance data to export")
            return
        
        try:
            # Filter data based on user role
            if self.user_role == 'admin':
                # Admin sees all data
                export_data = self.attendance_data
            else:
                # Roles User sees only their own data
                export_data = [record for record in self.attendance_data if record['user_id'] == self.current_user_id]
            
            if not export_data:
                messagebox.showwarning("Warning", "No attendance data to export")
                return
            
            # Create DataFrame
            df = pd.DataFrame(export_data)
            
            # Generate filename with timestamp
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            filename = f"wfh_attendance_{timestamp}.xlsx"
            
            # Save to Downloads folder
            if sys.platform == "win32":
                downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
            elif sys.platform == "darwin":
                downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
            else:
                downloads_path = os.path.join(os.path.expanduser("~"), "Downloads")
            
            filepath = os.path.join(downloads_path, filename)
            
            # Export to Excel
            df.to_excel(filepath, index=False, engine='openpyxl')
            
            # Update export path label
            self.export_path_var.set(f"Exported to: {filepath}")
            
            # Save export history
            self.save_export_history(filepath, len(export_data))
            
            # For Roles Users only: Clear their own records after export and start new session
            if self.user_role == 'roles':
                # Clear only the Roles User's records
                self.attendance_data = [record for record in self.attendance_data if record['user_id'] != self.current_user_id]
                self.save_data()
                
                # Show success message with record count
                messagebox.showinfo(
                    "Success", 
                    f"Data exported successfully!\n\n"
                    f"Exported {len(export_data)} records to:\n{filepath}\n\n"
                    f"Your attendance records have been cleared and new session started."
                )
                
                # Automatically create new session for Roles User
                if self.current_user_id and not self.current_session_id:
                    self.create_new_session()
            else:
                # Admin: Just show success message without clearing data
                messagebox.showinfo(
                    "Success", 
                    f"Data exported successfully!\n\n"
                    f"Exported {len(export_data)} records to:\n{filepath}"
                )
            
            # Update records display
            self.update_records_display()
            
            # Ask if user wants to open the file
            if messagebox.askyesno("Open File", "Do you want to open the Excel file?"):
                self.open_file(filepath)
                
        except Exception as e:
            messagebox.showerror("Error", f"Failed to export data: {str(e)}")
    
    def save_export_history(self, filepath: str, record_count: int):
        """Save export history for tracking"""
        export_record = {
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            'filepath': filepath,
            'record_count': record_count,
            'user_id': self.current_user_id,
            'user_role': self.user_role
        }
        
        self.export_history.append(export_record)
        
        try:
            with open(self.export_history_file, 'w') as f:
                json.dump(self.export_history, f, indent=2)
        except Exception as e:
            print(f"Error saving export history: {e}")
    
    def load_export_history(self) -> List[Dict]:
        """Load export history from JSON file"""
        try:
            if os.path.exists(self.export_history_file):
                with open(self.export_history_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading export history: {e}")
        return []
    
    def load_registered_users(self) -> List[Dict]:
        """Load registered users from JSON file"""
        try:
            if os.path.exists(self.users_file):
                with open(self.users_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading registered users: {e}")
        return []
    
    def save_registered_users(self):
        """Save registered users to JSON file"""
        try:
            with open(self.users_file, 'w') as f:
                json.dump(self.registered_users, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save user data: {str(e)}")
    
    def load_admin_users(self) -> List[Dict]:
        """Load admin users from JSON file"""
        try:
            if os.path.exists(self.admin_file):
                with open(self.admin_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading admin users: {e}")
        return []
    
    def save_admin_users(self):
        """Save admin users to JSON file"""
        try:
            with open(self.admin_file, 'w') as f:
                json.dump(self.admin_users, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save admin data: {str(e)}")
    
    def load_roles_users(self) -> List[Dict]:
        """Load roles users from JSON file"""
        try:
            if os.path.exists(self.roles_file):
                with open(self.roles_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading roles users: {e}")
        return []
    
    def save_roles_users(self):
        """Save roles users to JSON file"""
        try:
            with open(self.roles_file, 'w') as f:
                json.dump(self.roles_users, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save roles data: {str(e)}")
    
    def load_archive(self) -> List[Dict]:
        """Load deleted users archive from JSON file"""
        try:
            if os.path.exists(self.archive_file):
                with open(self.archive_file, 'r') as f:
                    return json.load(f)
        except Exception as e:
            print(f"Error loading archive: {e}")
        return []
    
    def save_archive(self):
        """Save deleted users archive to JSON file"""
        try:
            with open(self.archive_file, 'w') as f:
                json.dump(self.deleted_users_archive, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save archive: {str(e)}")
    
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
        """Update the records treeview based on user role"""
        # Clear existing records
        for item in self.records_tree.get_children():
            self.records_tree.delete(item)
        
        # Filter records based on user role
        if self.user_role == 'admin':
            # Admin sees all records
            display_data = self.attendance_data
        else:
            # Regular and Roles users see only their own records
            display_data = [record for record in self.attendance_data if record['user_id'] == self.current_user_id]
        
        # Add records (show latest first)
        for record in reversed(display_data):
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

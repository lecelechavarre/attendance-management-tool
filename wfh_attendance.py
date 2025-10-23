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
        self.root.title("WFH Attendance System")
        self.root.geometry("1400x900")
        self.root.configure(bg='#f8fafc')
        
        # Set window icon (you can add an icon file if needed)
        try:
            self.root.iconbitmap('attendance_icon.ico')
        except:
            pass
        
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
        
        # Configure modern professional styles
        self.setup_modern_styles()
        
        # Create modern UI
        self.create_modern_ui()
        
        # Update records display
        self.update_records_display()
        self.update_sessions_display()
        
        # Initially hide features based on role
        self.toggle_features_based_on_role()

    def center_window(self):
        """Center the window on the screen"""
        self.root.update_idletasks()
        width = 1400
        height = 900
        x = (self.root.winfo_screenwidth() // 2) - (width // 2)
        y = (self.root.winfo_screenheight() // 2) - (height // 2)
        self.root.geometry(f'{width}x{height}+{x}+{y}')

    def setup_modern_styles(self):
        """Configure modern, professional styles"""
        style = ttk.Style()
        style.theme_use('clam')
        
        # Professional color scheme
        self.colors = {
            'primary': '#2563eb',      # Professional blue
            'primary_light': '#3b82f6',
            'primary_dark': '#1d4ed8',
            'secondary': '#64748b',    # Slate gray
            'success': '#059669',      # Emerald
            'success_light': '#10b981',
            'warning': '#d97706',      # Amber
            'warning_light': '#f59e0b',
            'danger': '#dc2626',       # Red
            'danger_light': '#ef4444',
            'info': '#7c3aed',         # Violet
            'info_light': '#8b5cf6',
            'light': '#f8fafc',        # Light background
            'card_bg': '#ffffff',      # White cards
            'text_primary': '#1e293b', # Dark text
            'text_secondary': '#64748b', # Light text
            'text_light': '#94a3b8',   # Lighter text
            'border': '#e2e8f0',       # Border color
            'hover': '#f1f5f9'         # Hover state
        }
        
        # Configure styles
        style.configure('Modern.TFrame', background=self.colors['light'])
        style.configure('Card.TFrame', background=self.colors['card_bg'], relief='flat', borderwidth=1)
        
        # Title label
        style.configure('Title.TLabel',
                       font=('Segoe UI', 28, 'bold'),
                       background=self.colors['light'],
                       foreground=self.colors['text_primary'])
        
        # Section titles
        style.configure('Section.TLabel',
                       font=('Segoe UI', 16, 'bold'),
                       background=self.colors['card_bg'],
                       foreground=self.colors['text_primary'])
        
        # Regular labels
        style.configure('Modern.TLabel',
                       font=('Segoe UI', 11),
                       background=self.colors['card_bg'],
                       foreground=self.colors['text_primary'])
        
        # Small labels
        style.configure('Small.TLabel',
                       font=('Segoe UI', 9),
                       background=self.colors['card_bg'],
                       foreground=self.colors['text_secondary'])
        
        # Entry fields
        style.configure('Modern.TEntry',
                       font=('Segoe UI', 11),
                       fieldbackground='white',
                       borderwidth=2,
                       relief='solid',
                       focusthickness=3,
                       focuscolor=self.colors['primary'],
                       padding=(10, 8))
        
        # Buttons - Primary
        style.configure('Primary.TButton',
                       font=('Segoe UI', 11, 'bold'),
                       background=self.colors['primary'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none',
                       padding=(20, 12))
        style.map('Primary.TButton',
                 background=[('active', self.colors['primary_light']),
                           ('pressed', self.colors['primary_dark'])])
        
        # Success button
        style.configure('Success.TButton',
                       font=('Segoe UI', 11, 'bold'),
                       background=self.colors['success'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none',
                       padding=(20, 12))
        style.map('Success.TButton',
                 background=[('active', self.colors['success_light']),
                           ('pressed', '#047857')])
        
        # Warning button
        style.configure('Warning.TButton',
                       font=('Segoe UI', 11, 'bold'),
                       background=self.colors['warning'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none',
                       padding=(20, 12))
        style.map('Warning.TButton',
                 background=[('active', self.colors['warning_light']),
                           ('pressed', '#b45309')])
        
        # Danger button
        style.configure('Danger.TButton',
                       font=('Segoe UI', 11, 'bold'),
                       background=self.colors['danger'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none',
                       padding=(20, 12))
        style.map('Danger.TButton',
                 background=[('active', self.colors['danger_light']),
                           ('pressed', '#b91c1c')])
        
        # Info button
        style.configure('Info.TButton',
                       font=('Segoe UI', 11, 'bold'),
                       background=self.colors['info'],
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none',
                       padding=(20, 12))
        style.map('Info.TButton',
                 background=[('active', self.colors['info_light']),
                           ('pressed', '#6d28d9')])
        
        # Secondary button
        style.configure('Secondary.TButton',
                       font=('Segoe UI', 11),
                       background=self.colors['card_bg'],
                       foreground=self.colors['text_primary'],
                       borderwidth=1,
                       relief='solid',
                       focuscolor='none',
                       padding=(16, 10))
        style.map('Secondary.TButton',
                 background=[('active', self.colors['hover'])])
        
        # Treeview style
        style.configure('Modern.Treeview',
                       font=('Segoe UI', 10),
                       rowheight=35,
                       background='white',
                       fieldbackground='white',
                       foreground=self.colors['text_primary'],
                       borderwidth=0)
        style.configure('Modern.Treeview.Heading',
                       font=('Segoe UI', 11, 'bold'),
                       background=self.colors['primary'],
                       foreground='white',
                       relief='flat')
        style.map('Modern.Treeview.Heading',
                 background=[('active', self.colors['primary_light'])])
        
        # Notebook style
        style.configure('Modern.TNotebook',
                       background=self.colors['light'],
                       borderwidth=0)
        style.configure('Modern.TNotebook.Tab',
                       font=('Segoe UI', 11, 'bold'),
                       padding=(20, 12),
                       background=self.colors['card_bg'],
                       foreground=self.colors['text_secondary'],
                       borderwidth=1)
        style.map('Modern.TNotebook.Tab',
                 background=[('selected', self.colors['primary']),
                           ('active', self.colors['primary_light'])],
                 foreground=[('selected', 'white')])
        
        # Combobox style
        style.configure('Modern.TCombobox',
                       font=('Segoe UI', 10),
                       background='white',
                       fieldbackground='white',
                       borderwidth=2,
                       relief='solid')
        
        # Checkbutton style
        style.configure('Modern.TCheckbutton',
                       font=('Segoe UI', 10),
                       background=self.colors['card_bg'],
                       foreground=self.colors['text_primary'])

    def create_modern_ui(self):
        """Create modern, professional UI"""
        # Main container with soft background
        self.main_container = ttk.Frame(self.root, style='Modern.TFrame')
        self.main_container.pack(fill=tk.BOTH, expand=True, padx=0, pady=0)
        
        # Header section
        self.create_header()
        
        # Main content area
        self.create_content_area()

    def create_header(self):
        """Create modern header section"""
        header_frame = tk.Frame(self.main_container, bg=self.colors['primary'], height=100)
        header_frame.pack(fill=tk.X, padx=0, pady=0)
        header_frame.pack_propagate(False)
        
        # Header content
        header_content = ttk.Frame(header_frame, style='Modern.TFrame')
        header_content.place(relx=0.5, rely=0.5, anchor='center', width=1360, height=80)
        
        # Logo and title
        title_frame = ttk.Frame(header_content, style='Modern.TFrame')
        title_frame.pack(side=tk.LEFT, padx=40)
        
        # Main title with icon
        title_label = ttk.Label(
            title_frame,
            text="🏢 WFH Attendance System",
            style='Title.TLabel'
        )
        title_label.pack(side=tk.LEFT)
        
        # Time display
        time_frame = ttk.Frame(header_content, style='Modern.TFrame')
        time_frame.pack(side=tk.RIGHT, padx=40)
        
        self.time_display_var = tk.StringVar(value="Loading...")
        time_label = ttk.Label(
            time_frame,
            textvariable=self.time_display_var,
            font=('Segoe UI', 12, 'bold'),
            background=self.colors['primary'],
            foreground='white'
        )
        time_label.pack(side=tk.RIGHT)
        
        # Start clock update
        self.update_clock()

    def create_content_area(self):
        """Create main content area"""
        # Create notebook for main content
        self.main_notebook = ttk.Notebook(self.main_container, style='Modern.TNotebook')
        self.main_notebook.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Dashboard tab
        self.create_dashboard_tab()
        
        # Records tab
        self.create_records_tab()
        
        # Sessions tab
        self.create_sessions_tab()

    def create_dashboard_tab(self):
        """Create dashboard tab with login and attendance features"""
        dashboard_frame = ttk.Frame(self.main_notebook, style='Card.TFrame')
        self.main_notebook.add(dashboard_frame, text="📊 Dashboard")
        
        # Two-column layout
        main_content = ttk.Frame(dashboard_frame, style='Card.TFrame')
        main_content.pack(fill=tk.BOTH, expand=True, padx=40, pady=30)
        
        # Left column - Login and User Info
        left_column = ttk.Frame(main_content, style='Card.TFrame')
        left_column.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 20))
        
        # Right column - Attendance and Export
        right_column = ttk.Frame(main_content, style='Card.TFrame')
        right_column.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(20, 0))
        
        # Login Section
        self.create_login_section(left_column)
        
        # Attendance Section
        self.create_attendance_section(right_column)
        
        # Export Section
        self.create_export_section(right_column)

    def create_login_section(self, parent):
        """Create modern login section"""
        login_card = ttk.Frame(parent, style='Card.TFrame')
        login_card.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        # Card header
        header_frame = ttk.Frame(login_card, style='Card.TFrame')
        header_frame.pack(fill=tk.X, padx=25, pady=(25, 0))
        
        ttk.Label(
            header_frame,
            text="🔐 User Authentication",
            style='Section.TLabel'
        ).pack(anchor=tk.W)
        
        # Card content
        content_frame = ttk.Frame(login_card, style='Card.TFrame')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=20)
        
        # Form fields
        field_frame = ttk.Frame(content_frame, style='Card.TFrame')
        field_frame.pack(fill=tk.X, pady=(0, 20))
        
        # User ID
        ttk.Label(field_frame, text="User ID", style='Modern.TLabel').pack(anchor=tk.W, pady=(0, 8))
        self.user_id_var = tk.StringVar()
        self.user_id_entry = ttk.Entry(
            field_frame, 
            textvariable=self.user_id_var,
            font=('Segoe UI', 12),
            style='Modern.TEntry'
        )
        self.user_id_entry.pack(fill=tk.X, pady=(0, 20))
        
        # User Name
        ttk.Label(field_frame, text="User Name", style='Modern.TLabel').pack(anchor=tk.W, pady=(0, 8))
        self.user_name_var = tk.StringVar()
        self.user_name_entry = ttk.Entry(
            field_frame, 
            textvariable=self.user_name_var,
            font=('Segoe UI', 12),
            style='Modern.TEntry'
        )
        self.user_name_entry.pack(fill=tk.X, pady=(0, 25))
        
        # Buttons frame
        btn_frame = ttk.Frame(content_frame, style='Card.TFrame')
        btn_frame.pack(fill=tk.X)
        
        # Login Button
        self.login_btn = ttk.Button(
            btn_frame,
            text="🚀 Login",
            command=self.handle_login,
            style='Primary.TButton'
        )
        self.login_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Logout Button
        self.logout_btn = ttk.Button(
            btn_frame,
            text="🚪 Logout",
            command=self.handle_logout,
            style='Secondary.TButton',
            state=tk.DISABLED
        )
        self.logout_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Manage Users Button (Admin only)
        self.manage_users_btn = ttk.Button(
            btn_frame,
            text="👥 Manage Users",
            command=self.manage_users,
            style='Secondary.TButton',
            state=tk.DISABLED
        )
        self.manage_users_btn.pack(side=tk.LEFT)
        
        # Status Label
        status_frame = ttk.Frame(content_frame, style='Card.TFrame')
        status_frame.pack(fill=tk.X, pady=(20, 0))
        
        self.login_status_var = tk.StringVar(value="Please login to continue")
        self.login_status_label = ttk.Label(
            status_frame,
            textvariable=self.login_status_var,
            style='Small.TLabel'
        )
        self.login_status_label.pack(anchor=tk.W)

    def create_attendance_section(self, parent):
        """Create modern attendance section"""
        attendance_card = ttk.Frame(parent, style='Card.TFrame')
        attendance_card.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
        
        # Card header
        header_frame = ttk.Frame(attendance_card, style='Card.TFrame')
        header_frame.pack(fill=tk.X, padx=25, pady=(25, 0))
        
        ttk.Label(
            header_frame,
            text="⏰ Attendance Tracking",
            style='Section.TLabel'
        ).pack(anchor=tk.W)
        
        # Card content
        content_frame = ttk.Frame(attendance_card, style='Card.TFrame')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=20)
        
        # Status display
        status_frame = ttk.Frame(content_frame, style='Card.TFrame')
        status_frame.pack(fill=tk.X, pady=(0, 25))
        
        self.attendance_status_var = tk.StringVar(value="No active session")
        status_label = ttk.Label(
            status_frame,
            textvariable=self.attendance_status_var,
            font=('Segoe UI', 13, 'bold'),
            background=self.colors['card_bg'],
            foreground=self.colors['text_secondary']
        )
        status_label.pack(anchor=tk.W)
        
        # Buttons frame
        btn_frame = ttk.Frame(content_frame, style='Card.TFrame')
        btn_frame.pack(fill=tk.X)
        
        # Time In Button
        self.time_in_btn = ttk.Button(
            btn_frame,
            text="🟢 Time In",
            command=self.time_in,
            style='Success.TButton',
            state=tk.DISABLED
        )
        self.time_in_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Time Out Button
        self.time_out_btn = ttk.Button(
            btn_frame,
            text="🔴 Time Out",
            command=self.time_out,
            style='Danger.TButton',
            state=tk.DISABLED
        )
        self.time_out_btn.pack(side=tk.LEFT, padx=(0, 10))
        
        # Auto Time In Button (Roles User only)
        self.auto_time_in_btn = ttk.Button(
            btn_frame,
            text="🔄 Auto New Session",
            command=self.auto_new_session,
            style='Warning.TButton',
            state=tk.DISABLED
        )
        self.auto_time_in_btn.pack(side=tk.LEFT)

    def create_export_section(self, parent):
        """Create modern export section"""
        self.export_card = ttk.Frame(parent, style='Card.TFrame')
        
        # Card header
        header_frame = ttk.Frame(self.export_card, style='Card.TFrame')
        header_frame.pack(fill=tk.X, padx=25, pady=(25, 0))
        
        ttk.Label(
            header_frame,
            text="📤 Data Export",
            style='Section.TLabel'
        ).pack(anchor=tk.W)
        
        # Card content
        content_frame = ttk.Frame(self.export_card, style='Card.TFrame')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=20)
        
        # Export Button (Admin and Roles only)
        self.export_btn = ttk.Button(
            content_frame,
            text="💾 Export to Excel",
            command=self.export_to_excel,
            style='Info.TButton'
        )
        self.export_btn.pack(fill=tk.X)
        
        # Export Path Label
        path_frame = ttk.Frame(content_frame, style='Card.TFrame')
        path_frame.pack(fill=tk.X, pady=(15, 0))
        
        self.export_path_var = tk.StringVar(value="Export path will appear here")
        self.export_path_label = ttk.Label(
            path_frame,
            textvariable=self.export_path_var,
            style='Small.TLabel'
        )
        self.export_path_label.pack(anchor=tk.W)

    def create_records_tab(self):
        """Create modern records tab"""
        records_frame = ttk.Frame(self.main_notebook, style='Card.TFrame')
        self.main_notebook.add(records_frame, text="📋 Attendance Records")
        
        # Content with padding
        content_frame = ttk.Frame(records_frame, style='Card.TFrame')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        # Header
        header_frame = ttk.Frame(content_frame, style='Card.TFrame')
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(
            header_frame,
            text="My Attendance History",
            style='Section.TLabel'
        ).pack(anchor=tk.W)
        
        # Create treeview with scrollbar
        tree_frame = ttk.Frame(content_frame, style='Card.TFrame')
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
            style='Modern.Treeview',
            yscrollcommand=scrollbar.set
        )
        
        # Configure columns
        column_configs = [
            ('user_id', 'User ID', 120),
            ('user_name', 'User Name', 180),
            ('date', 'Date', 140),
            ('time_in', 'Time In', 120),
            ('time_out', 'Time Out', 120),
            ('duration', 'Duration', 120)
        ]
        
        for col, heading, width in column_configs:
            self.records_tree.heading(col, text=heading)
            self.records_tree.column(col, width=width, anchor=tk.CENTER)
        
        self.records_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.records_tree.yview)
        
        # Add alternating row colors
        self.records_tree.tag_configure('evenrow', background=self.colors['light'])
        self.records_tree.tag_configure('oddrow', background='white')

    def create_sessions_tab(self):
        """Create modern sessions tab"""
        sessions_frame = ttk.Frame(self.main_notebook, style='Card.TFrame')
        self.main_notebook.add(sessions_frame, text="🔍 Active Sessions")
        
        # Content with padding
        content_frame = ttk.Frame(sessions_frame, style='Card.TFrame')
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        # Header
        header_frame = ttk.Frame(content_frame, style='Card.TFrame')
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        ttk.Label(
            header_frame,
            text="Current Active Sessions",
            style='Section.TLabel'
        ).pack(anchor=tk.W)
        
        # Create treeview with scrollbar
        tree_frame = ttk.Frame(content_frame, style='Card.TFrame')
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
            style='Modern.Treeview',
            yscrollcommand=scrollbar.set
        )
        
        # Configure columns
        column_configs = [
            ('session_id', 'Session ID', 200),
            ('user_id', 'User ID', 120),
            ('user_name', 'User Name', 180),
            ('date', 'Date', 140),
            ('time_in', 'Time In', 120)
        ]
        
        for col, heading, width in column_configs:
            self.sessions_tree.heading(col, text=heading)
            self.sessions_tree.column(col, width=width, anchor=tk.CENTER)
        
        self.sessions_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.sessions_tree.yview)
        
        # Add alternating row colors
        self.sessions_tree.tag_configure('evenrow', background=self.colors['light'])
        self.sessions_tree.tag_configure('oddrow', background='white')
        
        # Force time out button (Admin only)
        self.force_out_btn = ttk.Button(
            content_frame,
            text="🛑 Force Time Out Selected Session",
            command=self.force_time_out,
            style='Danger.TButton'
        )

    def update_clock(self):
        """Update the current time display"""
        current_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.time_display_var.set(f"🕒 {current_time}")
        self.root.after(1000, self.update_clock)

    def toggle_features_based_on_role(self):
        """Show/hide features based on user role"""
        if self.user_role == 'admin':
            # Admin: Show user management, export, hide attendance
            self.export_card.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
            self.manage_users_btn.config(state=tk.NORMAL)
            self.force_out_btn.pack(pady=(20, 0))
            self.auto_time_in_btn.pack_forget()
        elif self.user_role == 'roles':
            # Roles User: Show export and attendance, hide user management
            self.export_card.pack(fill=tk.BOTH, expand=True, pady=(0, 20))
            self.manage_users_btn.config(state=tk.DISABLED)
            self.force_out_btn.pack_forget()
            self.auto_time_in_btn.pack(side=tk.LEFT, padx=(10, 0))
        elif self.user_role == 'regular':
            # Regular User: Show only attendance, hide everything else
            self.export_card.pack_forget()
            self.manage_users_btn.config(state=tk.DISABLED)
            self.force_out_btn.pack_forget()
            self.auto_time_in_btn.pack_forget()
        else:
            # Not logged in: Hide everything
            self.export_card.pack_forget()
            self.manage_users_btn.config(state=tk.DISABLED)
            self.force_out_btn.pack_forget()
            self.auto_time_in_btn.pack_forget()

    # ========== ALL YOUR ORIGINAL FUNCTIONAL METHODS - UNCHANGED ==========

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
            role_icon = "👑" if self.user_role == 'admin' else "⚡" if self.user_role == 'roles' else "👤"
            self.login_status_var.set(f"{role_icon} Logged in as: {user_name} ({user_id}) - {role_msg}")
            self.login_status_label.configure(foreground=self.colors['success'])
            
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
            
        # Create modern user management window
        users_window = tk.Toplevel(self.root)
        users_window.title("User Management - Admin")
        users_window.geometry("1000x700")
        users_window.configure(bg=self.colors['light'])
        
        # Center the window
        users_window.update_idletasks()
        width = 1000
        height = 700
        x = (users_window.winfo_screenwidth() // 2) - (width // 2)
        y = (users_window.winfo_screenheight() // 2) - (height // 2)
        users_window.geometry(f'{width}x{height}+{x}+{y}')
        
        # Main container
        main_container = ttk.Frame(users_window, style='Modern.TFrame')
        main_container.pack(fill=tk.BOTH, expand=True, padx=30, pady=30)
        
        # Title
        title_label = ttk.Label(
            main_container,
            text="👥 User Management - Admin Panel",
            style='Title.TLabel'
        )
        title_label.pack(pady=(0, 25))
        
        # Registration Card
        reg_card = ttk.Frame(main_container, style='Card.TFrame')
        reg_card.pack(fill=tk.X, pady=(0, 25))
        
        # Card header
        reg_header = ttk.Frame(reg_card, style='Card.TFrame')
        reg_header.pack(fill=tk.X, padx=25, pady=(20, 0))
        
        ttk.Label(
            reg_header,
            text="➕ Register New User",
            style='Section.TLabel'
        ).pack(anchor=tk.W)
        
        # Registration form
        form_frame = ttk.Frame(reg_card, style='Card.TFrame')
        form_frame.pack(fill=tk.X, padx=25, pady=20)
        
        # User ID
        ttk.Label(form_frame, text="User ID:", style='Modern.TLabel').grid(row=0, column=0, padx=(0, 15), pady=8, sticky=tk.W)
        new_user_id_var = tk.StringVar()
        new_user_id_entry = ttk.Entry(form_frame, textvariable=new_user_id_var, width=18, font=('Segoe UI', 11), style='Modern.TEntry')
        new_user_id_entry.grid(row=0, column=1, padx=(0, 30), pady=8)
        
        # User Name
        ttk.Label(form_frame, text="User Name:", style='Modern.TLabel').grid(row=0, column=2, padx=(0, 15), pady=8, sticky=tk.W)
        new_user_name_var = tk.StringVar()
        new_user_name_entry = ttk.Entry(form_frame, textvariable=new_user_name_var, width=25, font=('Segoe UI', 11), style='Modern.TEntry')
        new_user_name_entry.grid(row=0, column=3, padx=(0, 30), pady=8)
        
        # Role Selection
        ttk.Label(form_frame, text="Role:", style='Modern.TLabel').grid(row=0, column=4, padx=(0, 15), pady=8, sticky=tk.W)
        role_var = tk.StringVar(value="regular")
        role_combo = ttk.Combobox(form_frame, textvariable=role_var, values=["regular", "roles", "admin"], 
                                 state="readonly", width=12, style='Modern.TCombobox')
        role_combo.grid(row=0, column=5, padx=(0, 20), pady=8)
        
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
            
            # Show confirmation modal when registering admin
            if role == 'admin':
                confirm = messagebox.askyesno(
                    "Register Admin User",
                    f"Are you sure you want to register '{user_name}' ({user_id}) as an Administrator?\n\n"
                    f"Admin users have full system access including:\n"
                    f"• User management\n"
                    f"• Data export\n"
                    f"• Force session management\n\n"
                    f"This action cannot be undone."
                )
                if not confirm:
                    # Reset to regular user if not confirmed
                    role_var.set("regular")
                    return
            
            # Register new user
            self.register_new_user(user_id, user_name, role)
            
            # Show appropriate success message
            if role == 'admin':
                messagebox.showinfo("Success", f"Administrator '{user_name}' ({user_id}) registered successfully!")
            else:
                messagebox.showinfo("Success", f"User '{user_name}' ({user_id}) registered as {role} successfully!")
            
            # Clear fields but keep the selected role
            new_user_id_var.set("")
            new_user_name_var.set("")
            # Don't reset the role selection
            
            # Refresh user list
            refresh_user_list()
        
        register_btn = ttk.Button(form_frame, text="Register User", command=register_new_user_admin, style='Success.TButton')
        register_btn.grid(row=0, column=6, padx=(0, 10), pady=8)
        
        # Users List Card
        list_card = ttk.Frame(main_container, style='Card.TFrame')
        list_card.pack(fill=tk.BOTH, expand=True, pady=(0, 25))
        
        # Card header
        list_header = ttk.Frame(list_card, style='Card.TFrame')
        list_header.pack(fill=tk.X, padx=25, pady=(20, 0))
        
        ttk.Label(
            list_header,
            text="📋 Registered Users",
            style='Section.TLabel'
        ).pack(anchor=tk.W)
        
        # Create treeview with scrollbar
        tree_frame = ttk.Frame(list_card, style='Card.TFrame')
        tree_frame.pack(fill=tk.BOTH, expand=True, padx=25, pady=20)
        
        scrollbar = ttk.Scrollbar(tree_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Treeview
        columns = ('user_id', 'user_name', 'role', 'registered_date')
        users_tree = ttk.Treeview(
            tree_frame,
            columns=columns,
            show='headings',
            style='Modern.Treeview',
            yscrollcommand=scrollbar.set
        )
        
        # Configure columns
        column_configs = [
            ('user_id', 'User ID', 120),
            ('user_name', 'User Name', 180),
            ('role', 'Role', 100),
            ('registered_date', 'Registration Date', 180)
        ]
        
        for col, heading, width in column_configs:
            users_tree.heading(col, text=heading)
            users_tree.column(col, width=width, anchor=tk.CENTER)
        
        users_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=users_tree.yview)
        
        # Add alternating row colors
        users_tree.tag_configure('evenrow', background=self.colors['light'])
        users_tree.tag_configure('oddrow', background='white')
        
        def refresh_user_list():
            # Clear existing items
            for item in users_tree.get_children():
                users_tree.delete(item)
            
            # Add users to treeview with alternating colors
            for i, user in enumerate(self.registered_users):
                tag = 'evenrow' if i % 2 == 0 else 'oddrow'
                users_tree.insert('', tk.END, values=(
                    user['user_id'],
                    user['user_name'],
                    user.get('role', 'regular'),
                    user['registered_date']
                ), tags=(tag,))
        
        def delete_selected_user():
            selected = users_tree.selection()
            if not selected:
                messagebox.showwarning("Warning", "Please select a user to delete")
                return
            
            item = selected[0]
            user_id = users_tree.item(item, 'values')[0]
            user_name = users_tree.item(item, 'values')[1]
            user_role = users_tree.item(item, 'values')[2]
            
            # Prevent self-deletion
            if user_id == self.current_user_id:
                messagebox.showerror("Error", "You cannot delete your own account")
                return
            
            # Extra confirmation for deleting admin users
            if user_role == 'admin':
                confirm = messagebox.askyesno(
                    "Confirm Admin Deletion",
                    f"WARNING: You are about to delete an Administrator account!\n\n"
                    f"User: '{user_name}' ({user_id})\n\n"
                    f"This action will remove all admin privileges from this user.\n"
                    f"Are you absolutely sure you want to proceed?"
                )
            else:
                confirm = messagebox.askyesno(
                    "Confirm Deletion", 
                    f"Are you sure you want to delete user '{user_name}' ({user_id})?"
                )
                
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
        btn_frame = ttk.Frame(list_card, style='Card.TFrame')
        btn_frame.pack(fill=tk.X, padx=25, pady=(0, 20))
        
        delete_btn = ttk.Button(btn_frame, text="🗑️ Delete Selected User", command=delete_selected_user, style='Danger.TButton')
        delete_btn.pack(side=tk.LEFT, padx=(0, 15))
        
        refresh_btn = ttk.Button(btn_frame, text="🔄 Refresh List", command=refresh_user_list, style='Primary.TButton')
        refresh_btn.pack(side=tk.LEFT)
        
        close_btn = ttk.Button(main_container, text="Close", command=users_window.destroy, style='Secondary.TButton')
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
        self.login_status_label.configure(foreground=self.colors['text_secondary'])
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
            self.attendance_status_var.set(f"🟢 Active session: Time In at {session['time_in']}")
            self.time_in_btn.config(state=tk.DISABLED)
            self.time_out_btn.config(state=tk.NORMAL)
            self.auto_time_in_btn.config(state=tk.DISABLED)
        else:
            self.current_session_id = None
            self.attendance_status_var.set("🟡 Ready for Time In")
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
        
        self.attendance_status_var.set(f"🟢 Time In recorded at {current_time}")
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
        
        self.attendance_status_var.set(f"🔴 Time Out recorded at {current_time}")
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
            self.export_path_var.set(f"📁 Exported to: {filepath}")
            
            # Save export history
            self.save_export_history(filepath, len(export_data))
            
            # For Roles Users only - create a COPY for export, don't modify main data
            if self.user_role == 'roles':
                # Create a temporary copy of the user's data for display purposes only
                user_export_data = [record for record in self.attendance_data if record['user_id'] == self.current_user_id]
                
                # Show success message with record count
                messagebox.showinfo(
                    "Success", 
                    f"Data exported successfully!\n\n"
                    f"Exported {len(user_export_data)} records to:\n{filepath}\n\n"
                    f"Your data has been exported. The main attendance records remain unchanged."
                )
                
                # Automatically create new session for Roles User without clearing data
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
        
        # Add records (show latest first) with alternating colors
        for i, record in enumerate(reversed(display_data)):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            self.records_tree.insert(
                '', tk.END,
                values=(
                    record['user_id'],
                    record['user_name'],
                    record['date'],
                    record['time_in'],
                    record['time_out'],
                    record['duration']
                ),
                tags=(tag,)
            )

    def update_sessions_display(self):
        """Update the active sessions treeview"""
        # Clear existing sessions
        for item in self.sessions_tree.get_children():
            self.sessions_tree.delete(item)
        
        # Add active sessions with alternating colors
        for i, session in enumerate(self.active_sessions):
            tag = 'evenrow' if i % 2 == 0 else 'oddrow'
            self.sessions_tree.insert(
                '', tk.END,
                values=(
                    session['session_id'],
                    session['user_id'],
                    session['user_name'],
                    session['date'],
                    session['time_in']
                ),
                tags=(tag,)
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

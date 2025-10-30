WFH Attendance System

A desktop application for tracking Work From Home attendance with secure role-based access control and automated data export features.

Features
**User Authentication & Role Management**
- Three User Roles: Admin, Roles User, and Regular User
- Secure Login: User ID and Name verification
- Role-based Access Control: Different features available based on user role

**Attendance Tracking**
- Time In/Time Out: Record attendance with session management
- Active Session Monitoring: Real-time tracking of current sessions
- Auto Sessions: Roles users can create new sessions without time out
- Session Validation: Ensures users can only time out their own sessions

**Data Management**
- Time Records: View attendance history with filtering by user role
- Active Sessions Display: Monitor all current active sessions
- Data Export: Roles users can export attendance data to Excel
- Automatic Data Clearance: Records are cleared after export to maintain system efficiency

**Administrative Features**
- User Management: Admin users can register and manage all users
- Force Time Out: Admin can manually end any active session
- Export Access: Admin can download Excel files exported by roles users
- Read-only Protection: Exported files are read-only for roles users but editable for admin

**Prerequisites**
Python 3.7 or higher
Required Python packages (install via pip):

bash
pip install tkinter pandas openpyxl


🛠️ Technical Details
File Structure
wfh_attendance_system/
├── wfh_attendance_system.py  # Main application
├── attendance_data.json      # Attendance records
├── registered_users.json     # User database
├── active_sessions.json      # Current sessions
├── admin_users.json          # Admin users
├── roles_users.json          # Roles users
├── export_history.json       # Export log
├── deleted_users_archive.json # User archive
└── roles_exports/            # Admin-accessible exports

**Dependencies**
- tkinter: GUI framework
- pandas: Data manipulation and Excel export
- openpyxl: Excel file handling
- datetime: Time tracking and session management

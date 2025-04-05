# Todo Application Documentation

## Overview

The Todo Application is a feature-rich desktop application built using Visual Basic 4.0. It provides task management, team collaboration, and advanced features like recurring tasks, comments, and analytics.

---

## Project Structure

### Forms
1. **frmTodo.frm**  
   - Main interface for managing tasks.
   - Features task creation, editing, deletion, and filtering.

2. **frmComments.frm**  
   - Allows users to view and add comments to tasks.
   - Supports @mentions for team collaboration.

3. **frmTrash.frm**  
   - Displays deleted tasks.
   - Allows restoring or permanently deleting tasks.

4. **frmLogin.frm**  
   - Login form for user authentication.
   - Provides access to the registration form.

5. **frmRegister.frm**  
   - Registration form for new users.
   - Validates email, password, and phone number.

6. **frmEditTask.frm**  
   - Enables editing of task details like priority, category, and due date.

7. **frmDashboard.frm**  
   - Displays task statistics and progress.
   - Includes priority breakdown, category distribution, and pinned tasks.

8. **frmTeam.frm**  
   - Manages teams and team members.
   - Allows creating teams and assigning roles.

9. **frmTaskDetails.frm**  
   - Provides detailed view of a task.
   - Includes subtasks, notes, recurrence, and attachments.

---

### Modules
1. **modTodo.bas**  
   - Contains utility functions for string manipulation, date formatting, and UI adjustments.

2. **modDatabase.bas**  
   - Handles database initialization and operations.
   - Includes functions for saving, loading, and managing tasks, categories, and teams.

3. **modNotifications.bas**  
   - Manages email and SMS notifications.
   - Includes functions for sending alerts and reminders.

---

### Database
- **Database File:** `Todo.mdb`
- **Tables:**
  - `Users`: Stores user information.
  - `Tasks`: Stores task details.
  - `Comments`: Stores task comments.
  - `Mentions`: Tracks @mentions in comments.
  - `Categories`: Stores task categories.
  - `Teams`: Stores team information.
  - `TeamMembers`: Tracks team membership.
  - `Attachments`: Stores file attachments for tasks.
  - `DeletedTasks`: Tracks deleted tasks for undo functionality.

---

## Features

### Task Management
- Add, edit, delete, and restore tasks.
- Set task priority, category, due date, and recurrence.
- Pin important tasks for quick access.

### Team Collaboration
- Create and manage teams.
- Share tasks with team members.
- Add comments and @mention team members.

### Analytics
- View task completion progress.
- Analyze tasks by priority and category.
- Track pinned tasks and recurring tasks.

### Notifications
- Email and SMS reminders for due tasks.
- Notifications for @mentions in comments.

### Advanced Features
- Subtasks with independent tracking.
- Attach files to tasks.
- Export tasks to CSV format.

---

## How to Use

### Getting Started
1. Register a new account using the **frmRegister** form.
2. Login using the **frmLogin** form.
3. Start managing tasks in the **frmTodo** form.

### Task Management
1. Add tasks using the input fields in **frmTodo**.
2. Edit tasks by selecting a task and clicking "Edit."
3. Delete tasks and restore them from the **frmTrash** form.

### Team Collaboration
1. Create teams in the **frmTeam** form.
2. Add members to teams and assign roles.
3. Share tasks with team members and collaborate using comments.

### Analytics
1. Open the **frmDashboard** form to view task statistics.
2. Analyze task progress, priority distribution, and category breakdown.

---

## Configuration

### Email Notifications
- Configure SMTP settings in `modNotifications.bas`:
  ```vb
  Private Const SMTP_SERVER As String = "smtp.gmail.com"
  Private Const SMTP_PORT As Integer = 587
  Private Const SMTP_USERNAME As String = "your_email@gmail.com"
  Private Const SMTP_PASSWORD As String = "your_app_password"
  ```

### SMS Notifications
- Configure Twilio settings in `modNotifications.bas`:
  ```vb
  Private Const TWILIO_ACCOUNT_SID As String = "your_account_sid"
  Private Const TWILIO_AUTH_TOKEN As String = "your_auth_token"
  Private Const TWILIO_FROM_NUMBER As String = "your_twilio_number"
  ```

---

## Troubleshooting

1. **Application Fails to Start**
   - Ensure `VBRun400.dll` is installed.
   - Verify all required files are in the same directory.

2. **Database Errors**
   - Check if `Todo.mdb` exists in the application directory.
   - Ensure the application has write permissions.

3. **Email/SMS Notifications Not Working**
   - Verify SMTP/Twilio credentials in `modNotifications.bas`.
   - Check internet connectivity.

---

## Support

For any issues or questions, contact the development team or refer to the README file for additional instructions.

# Todo Application

A feature-rich desktop Todo application built with Visual Basic 4.0.

## System Requirements

- Microsoft Windows
- Visual Basic 4.0 Runtime (VBRun400.dll)
- Screen resolution: 800x600 or higher

## Installation

1. Copy all the following files to a directory on your computer:
   - TodoApp.vbp (Project file)
   - frmTodo.frm (Main form)
   - modTodo.bas (Module with utilities)
   - todos.txt (Will be created automatically when you save tasks)

## Running the Application

### Method 1: Using Visual Basic 4.0 IDE
1. Open Visual Basic 4.0
2. File -> Open Project
3. Navigate to the directory containing the files
4. Select `TodoApp.vbp`
5. Click Run (F5) or the Play button

### Method 2: Using Compiled Executable
1. Double-click the `TodoApp.exe` (after compiling the project)
2. The application will start automatically

## Features

1. **User Management**
   - User registration and login
   - Secure password storage
   - Personal task lists for each user
   - Email and phone number storage
   - Theme preferences (Light/Dark mode)

2. **Task Management**
   - Add new tasks with descriptions
   - Edit existing tasks
   - Mark tasks as complete/incomplete
   - Delete tasks with undo capability
   - Trash bin for deleted tasks
   - Pin important tasks to top
   - Task weights (1-5) for prioritization

3. **Task Properties**
   - Priority levels (High ⚡, Medium ●, Low ○)
   - Categories (Personal, Work, Shopping)
   - Due dates with reminders
   - Creation and deletion dates
   - Notes and rich text formatting
   - File attachments

4. **Subtasks & Hierarchy**
   - Create subtasks for complex items
   - Nested task hierarchy
   - Independent completion tracking
   - Inherit parent task properties
   - Visual indentation in lists

5. **Recurring Tasks**
   - Set tasks to repeat automatically
   - Daily, Weekly, Monthly, Yearly intervals
   - Customizable repeat frequency
   - Automatic task creation
   - Due date adjustment

6. **Team Collaboration**
   - Create and manage teams
   - Add team members with roles
   - Share tasks with team members
   - Team-specific task views
   - Team member notifications

7. **Comments & Mentions**
   - Add comments to tasks
   - @mention team members
   - Email notifications for mentions
   - Comment history with timestamps
   - Rich text formatting in comments

8. **Export & Sharing**
   - Export tasks to CSV format
   - Share tasks via unique links
   - Set expiry for shared links
   - Track share link usage
   - Share with non-team members

9. **Dashboard & Analytics**
   - Visual progress tracking
   - Task completion statistics
   - Priority distribution
   - Category breakdown
   - Team performance metrics

10. **Task Organization**
    - Sort by multiple criteria
    - Group by priority/category/date
    - Filter by various properties
    - Search task content
    - Custom task ordering

11. **UI/UX Features**
    - Dark/Light theme support
    - Modern visual indicators
    - Intuitive task hierarchy
    - Responsive interface
    - Keyboard shortcuts

12. **Data Management**
    - Automatic data saving
    - File attachment storage
    - Database relationships
    - Data integrity checks
    - Backup capabilities

## How to Use

1. **Getting Started**
   - Register a new account with email and phone
   - Login with your credentials
   - Choose your preferred theme (Light/Dark)
   - Your tasks are private and secure

2. **Task Creation**
   - Enter task description
   - Set priority (High ⚡, Medium ●, Low ○)
   - Choose category
   - Set due date (optional)
   - Add weight (1-5) for importance
   - Set recurrence (if needed)

3. **Task Details**
   - Click "Details" to open task details
   - Add/manage subtasks
   - Write notes with formatting
   - Attach relevant files
   - View task history

4. **Recurring Tasks**
   - Open task details
   - Choose recurrence type
   - Set interval (e.g., every 2 weeks)
   - System auto-creates new instances
   - Maintains task properties

5. **Task Organization**
   - Use "Sort by" to order tasks
   - Use "Group by" to categorize
   - Pin important tasks
   - Adjust task weights
   - Search and filter as needed

6. **Team Features**
   - Create/join teams
   - Share tasks with team
   - Add comments and mentions
   - Track team progress
   - Manage permissions

7. **Dashboard**
   - View completion rates
   - Check priority distribution
   - Monitor team performance
   - Track recurring tasks
   - Export statistics

8. **Customization**
   - Toggle Light/Dark theme
   - Configure notifications
   - Set default views
   - Customize categories
   - Adjust weights

## Task Display Format

Tasks are displayed in the following format:
`[Status] [Priority] [Category] Task Description (Due: Date)`

Example:
`[✓] [!] [Work] Complete project report (Due: 15/04/2025)`

Where:
- [✓] indicates completed task
- [!] indicates high priority
- [-] indicates medium priority
- [.] indicates low priority

## Configuration

1. **Email Settings**
   - Configure SMTP server in modNotifications.bas
   - Set your email credentials
   - Enable less secure app access if using Gmail

2. **SMS Settings**
   - Set up Twilio account
   - Configure account SID and auth token
   - Set your Twilio phone number

## Security Notes

1. Store sensitive information securely
2. Use strong passwords
3. Never share your login credentials
4. Regularly backup the database file

## Troubleshooting

1. If the application doesn't start:
   - Verify Visual Basic 4.0 Runtime is installed
   - Check if all files are in the same directory

2. If saving/loading fails:
   - Ensure you have write permissions in the application directory
   - Check if todos.txt is not read-only

## Support

For any issues or questions, please check:
1. File permissions
2. Runtime installation
3. System requirements

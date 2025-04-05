Feature: Comprehensive Scenarios for Todo Application

  # User Registration
  Scenario: Registering with valid details
    Given I am on the registration page
    When I enter a valid email, password, and phone number
    And I confirm my password
    And I click "Register"
    Then I should see a message "Registration successful!"

  Scenario: Registering with an already registered email
    Given I am on the registration page
    When I enter an email that is already registered
    And I click "Register"
    Then I should see an error message "This email is already registered!"

  Scenario: Registering with mismatched passwords
    Given I am on the registration page
    When I enter a password and a different confirm password
    And I click "Register"
    Then I should see an error message "Passwords do not match!"

  Scenario: Registering with an invalid email format
    Given I am on the registration page
    When I enter an invalid email format
    And I click "Register"
    Then I should see an error message "Please enter a valid email address!"

  Scenario: Registering with a short password
    Given I am on the registration page
    When I enter a password shorter than 6 characters
    And I click "Register"
    Then I should see an error message "Password must be at least 6 characters long!"

  # User Login
  Scenario: Logging in with valid credentials
    Given I am on the login page
    When I enter a valid email and password
    And I click "Login"
    Then I should be redirected to the main task page

  Scenario: Logging in with an unregistered email
    Given I am on the login page
    When I enter an unregistered email and a password
    And I click "Login"
    Then I should see an error message "Invalid email or password!"

  Scenario: Logging in with an incorrect password
    Given I am on the login page
    When I enter a valid email and an incorrect password
    And I click "Login"
    Then I should see an error message "Invalid email or password!"

  # Task Management
  Scenario: Adding a task with all valid details
    Given I am on the main task page
    When I enter a task description, select a priority, category, and due date
    And I click "Add Task"
    Then the task should appear in the task list

  Scenario: Adding a task without a description
    Given I am on the main task page
    When I leave the task description blank
    And I click "Add Task"
    Then I should see an error message "Please enter a task!"

  Scenario: Adding a task with an invalid due date
    Given I am on the main task page
    When I enter an invalid due date
    And I click "Add Task"
    Then I should see an error message "Please enter a valid date in dd/mm/yyyy format!"

  Scenario: Editing a task with valid details
    Given I have a task in the task list
    When I select the task
    And I update the task description, priority, category, and due date
    And I click "Save"
    Then the task should be updated in the task list

  Scenario: Deleting a task
    Given I have a task in the task list
    When I select the task
    And I click "Delete"
    Then the task should move to the trash

  Scenario: Restoring a task from the trash
    Given I have a task in the trash
    When I select the task
    And I click "Restore"
    Then the task should reappear in the task list

  Scenario: Permanently deleting a task from the trash
    Given I have a task in the trash
    When I select the task
    And I click "Delete Forever"
    Then the task should be permanently removed

  # Subtasks
  Scenario: Adding a subtask to a task
    Given I am viewing a task's details
    When I enter a subtask description
    And I click "Add"
    Then the subtask should appear under the task

  Scenario: Adding a subtask without a description
    Given I am viewing a task's details
    When I leave the subtask description blank
    And I click "Add"
    Then I should see an error message "Please enter a subtask description!"

  # Recurring Tasks
  Scenario: Setting a valid recurring task
    Given I am viewing a task's details
    When I select a recurrence type and interval
    And I click "Set"
    Then the task should be marked as recurring

  Scenario: Setting a recurring task with an invalid interval
    Given I am viewing a task's details
    When I enter a non-numeric interval
    And I click "Set"
    Then I should see an error message "Please enter a valid interval!"

  # Team Management
  Scenario: Creating a new team
    Given I am on the team management page
    When I enter a team name
    And I click "Create Team"
    Then the team should appear in my team list

  Scenario: Adding a team member with a valid email
    Given I have selected a team
    When I enter a valid member's email and select a role
    And I click "Add Member"
    Then the member should appear in the team member list

  Scenario: Adding a team member with an invalid email
    Given I have selected a team
    When I enter an invalid email
    And I click "Add Member"
    Then I should see an error message "User not found with this email!"

  Scenario: Adding a duplicate team member
    Given I have selected a team
    When I enter an email of an existing team member
    And I click "Add Member"
    Then I should see an error message "This user is already a team member!"

  # Comments
  Scenario: Adding a comment to a task
    Given I am viewing a task's comments
    When I enter a comment
    And I click "Add Comment"
    Then the comment should appear in the comment list

  Scenario: Adding a comment with @mentions
    Given I am viewing a task's comments
    When I enter a comment with "@member_email"
    And I click "Add Comment"
    Then the mentioned member should receive a notification

  Scenario: Adding a blank comment
    Given I am viewing a task's comments
    When I leave the comment field blank
    And I click "Add Comment"
    Then I should see an error message "Please enter a comment!"

  # Dashboard
  Scenario: Viewing task progress
    Given I am on the dashboard page
    When I view the progress overview
    Then I should see the percentage of completed tasks

  Scenario: Viewing tasks by priority
    Given I am on the dashboard page
    When I view the priority breakdown
    Then I should see the count of tasks for each priority level

  Scenario: Viewing pinned tasks
    Given I am on the dashboard page
    When I view the pinned tasks section
    Then I should see all pinned tasks

  # Notifications
  Scenario: Receiving a due date reminder
    Given I have a task due in 2 days
    When the reminder timer triggers
    Then I should receive an email and SMS notification

  Scenario: Receiving a notification for @mentions
    Given I am mentioned in a comment
    When the comment is added
    Then I should receive a notification

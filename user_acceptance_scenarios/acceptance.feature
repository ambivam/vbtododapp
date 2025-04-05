Feature: User Acceptance Scenarios

  Scenario: User registration and login
    Given I am on the registration page
    When I register with a valid email, password, and phone number
    Then I should see a message "Registration successful!"
    When I log in with the same email and password
    Then I should be redirected to the main task page

  Scenario: Task management workflow
    Given I am logged in as a user
    When I add a new task with a description, priority, and due date
    And I edit the task to update its details
    And I mark the task as complete
    Then the task should appear as completed in my task list

  Scenario: Team collaboration workflow
    Given I am logged in as a user
    When I create a new team
    And I add a team member
    And I share a task with the team
    Then the team member should see the task in their task list

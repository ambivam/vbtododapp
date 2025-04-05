Feature: End-to-End Scenarios

  Scenario: Registering, logging in, and adding a task
    Given I am on the registration page
    When I register with a valid email, password, and phone number
    And I log in with the same email and password
    And I add a new task with a description, priority, and due date
    Then the task should appear in my task list

  Scenario: Deleting a task and restoring it from the trash
    Given I have a task in my task list
    When I delete the task
    And I navigate to the trash
    And I restore the task
    Then the task should reappear in my task list

  Scenario: Setting a recurring task and viewing its next instance
    Given I have a task in my task list
    When I set the task to repeat weekly
    And the next recurrence date is reached
    Then a new instance of the task should appear in my task list

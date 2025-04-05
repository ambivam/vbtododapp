Feature: Task Management

  Scenario: Adding a new task
    Given I am on the main task page
    When I enter a task description
    And I select a priority and category
    And I set a due date
    And I click "Add Task"
    Then the task should appear in the task list

  Scenario: Editing an existing task
    Given I have a task in the task list
    When I select the task
    And I click "Edit"
    And I update the task details
    And I click "Save"
    Then the task should be updated in the task list

  Scenario: Deleting a task
    Given I have a task in the task list
    When I select the task
    And I click "Delete"
    Then the task should move to the trash

  Scenario: Marking a task as complete
    Given I have a task in the task list
    When I select the task
    And I click "Mark Complete"
    Then the task should be marked as completed

  Scenario: Pinning a task
    Given I have a task in the task list
    When I select the task
    And I click "Pin"
    Then the task should appear at the top of the task list

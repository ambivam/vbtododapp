Feature: Task Editing

  Scenario: Editing a task
    Given I have selected a task to edit
    When I update the task description, priority, category, and due date
    And I click "Save"
    Then the task should be updated in the task list

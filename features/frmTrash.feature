Feature: Trash Management

  Scenario: Restoring a task from trash
    Given I have a task in the trash
    When I select the task
    And I click "Restore"
    Then the task should reappear in the task list

  Scenario: Permanently deleting a task
    Given I have a task in the trash
    When I select the task
    And I click "Delete Forever"
    Then the task should be permanently removed

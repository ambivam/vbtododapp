Feature: Task Details

  Scenario: Adding a subtask
    Given I am viewing a task's details
    When I enter a subtask description
    And I click "Add"
    Then the subtask should appear under the task

  Scenario: Adding notes to a task
    Given I am viewing a task's details
    When I enter notes in the notes section
    And I click "Save"
    Then the notes should be saved for the task

  Scenario: Setting task recurrence
    Given I am viewing a task's details
    When I select a recurrence type and interval
    And I click "Set"
    Then the task should be marked as recurring

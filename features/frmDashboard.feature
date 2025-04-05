Feature: Task Dashboard

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

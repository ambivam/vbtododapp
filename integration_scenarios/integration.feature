Feature: Integration Scenarios

  Scenario: Adding a task and viewing it in the dashboard
    Given I am logged in as a user
    When I add a new task with a description, priority, and due date
    And I navigate to the dashboard
    Then I should see the task in the progress overview

  Scenario: Adding a comment and notifying a team member
    Given I am viewing a task's details
    When I add a comment with "@team_member_email"
    Then the mentioned team member should receive a notification

  Scenario: Creating a team and assigning tasks
    Given I am logged in as a user
    When I create a new team
    And I add a team member
    And I assign a task to the team
    Then the team member should see the task in their task list

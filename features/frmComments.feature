Feature: Task Comments

  Scenario: Adding a comment
    Given I am viewing a task's comments
    When I enter a comment
    And I click "Add Comment"
    Then the comment should appear in the comment list

  Scenario: Mentioning a team member
    Given I am adding a comment
    When I include "@member_email" in the comment
    And I click "Add Comment"
    Then the mentioned member should receive a notification

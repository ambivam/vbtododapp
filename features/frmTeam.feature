Feature: Team Management

  Scenario: Creating a new team
    Given I am on the team management page
    When I enter a team name
    And I click "Create Team"
    Then the team should appear in my team list

  Scenario: Adding a team member
    Given I have selected a team
    When I enter a member's email and select a role
    And I click "Add Member"
    Then the member should appear in the team member list

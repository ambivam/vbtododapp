Feature: User Login

  Scenario: Logging in with valid credentials
    Given I am on the login page
    When I enter a valid email and password
    And I click "Login"
    Then I should be redirected to the main task page

  Scenario: Registering a new account
    Given I am on the login page
    When I click "Register"
    Then I should be redirected to the registration page

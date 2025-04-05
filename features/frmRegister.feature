Feature: User Registration

  Scenario: Registering a new user
    Given I am on the registration page
    When I enter a valid email, password, and phone number
    And I confirm my password
    And I click "Register"
    Then I should see a message "Registration successful!"

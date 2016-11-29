Feature: Go to google

Scenario: Visit Google
  Given I go to "http://www.google.com/"
  When I fill in field with class "gsfi" with "testingbot"
  Then I should see "testingbot.com" within 2 second
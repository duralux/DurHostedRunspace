# DurHostedRunspace

Let's you easily acccess a powershell runspace via dependency injection.

Knwon issues:
 - Runspace pool is not running powershell script if the the script was called via asp.net core controller. Works well with runspace or runspace pools work well with any other (tested) method of running.
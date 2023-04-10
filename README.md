# quicksight-slack
Script that given an embeded dashboard URL to a quicksight table will print a message on Slack with all the information of that given table<br><br>

Arguments needed for the .jar to work are as follows:<br>
1. AWS Access Key ID (https://docs.aws.amazon.com/powershell/latest/userguide/pstools-appendix-sign-up.html)<br>
2. AWS Secret Access Key (https://docs.aws.amazon.com/powershell/latest/userguide/pstools-appendix-sign-up.html)<br>
3. AWS Region (found on the URL of the dashboard. Region where the dashboard is allocated with format: US-EAST-1)<br>
4. Dashboard ID (when you are on a Quicksight Dashboard generate a new Embed URL)<br>
5. Slack bot token (https://api.slack.com/authentication/token-types)<br>
6. The channel ID on slack to send the dashboard<br>

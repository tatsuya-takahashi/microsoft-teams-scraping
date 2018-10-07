# microsoft-teams-scraping
## Waht's this?
- This is the Scraping Module for Microsoft Teams with Grph API.
- this works on Windows, Linux, OS X.
- Written by node.js. You must intall npm and node.

## How To Use
### Register Your App to Active Directory
- https://qiita.com/tatsuya-takahashi/items/f0cfc2a00b5c9b885bcc

### Run App
- You Must Set below var Or Replace Code in index.js.  
  - var loginUrl = process.env.loginUrl;  // Your Login URL
  - var ADClientId = process.env.clientId;  // AD Client Id
  - var ADSecret = process.env.secret;  // AD Secret
- `npm install`
- `node index.js`
- and access to http://localhost:3000
- you'll get teams.tsv.

## Format
- i.e.  

|teamId|teamName|channelId|channelName|userId|name|content|
|---|---|---|---|---|---|---|
|9d93a0d9-c462-4cbb-9cb8-9badb0b65f45|ManagerTeam|05:6d03b62a324042e6885d2f8bf2d80b6e@thread.skype|General|b0543ed0-a012-4653-bbd7-9fe60c967af7|Smith|Hello, everyone.|
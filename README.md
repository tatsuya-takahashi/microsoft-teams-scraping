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

teams.tsv
|teamId|teamName|channelId|channelName|userId|name|content|
|---|---|---|---|---|---|---|
|9d93a0d9-c462-4cbb-9cb8-9badb0b65f45|ManagerTeam|05:6d03b62a324042e6885d2f8bf2d80b6e@thread.skype|General|b0543ed0-a012-4653-bbd7-9fe60c967af7|Smith|Hello, everyone.|

teams_relations
|relationType|fromUserId|toUserId|
|---|---|---|
|like|7de1cc33-077b-42c3-9fdd-2d5c62a31522|95eecce7-f85e-4c4b-9df3-4a7e1d02d03b|
|like|d9db64c1-bd8c-45a2-8552-012e54ebcb20|bde2bf3e-5307-416b-b469-d2c730f2ab65|
|reply|4a543f69-917d-4eb5-8f56-14955fde3ea1|bde2bf3e-5307-416b-b469-d2c730f2ab65|
|like|bde2bf3e-5307-416b-b469-d2c730f2ab65|4a543f69-917d-4eb5-8f56-14955fde3ea1|
|reply|7de1cc33-077b-42c3-9fdd-2d5c62a31522|5095ec96-e9be-492a-a8e7-7b1ce466a5dc|
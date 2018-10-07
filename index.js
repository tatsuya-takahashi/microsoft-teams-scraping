/**
 * Teams Scraping Apps
 */

// vars
var loginUrl = process.env.loginUrl;
var ADClientId = process.env.clientId;
var ADSecret = process.env.secret;

// import
const express = require("express");
const bodyParser = require("body-parser");
const request = require("request");
var app = express();
var access_token = "";
var messageText = [];
var fs = require('fs');

// Body Parser
app.use(bodyParser.json());

// Allow CORS
app.use((req, res, next) => {
    res.header('Access-Control-Allow-Origin', '*');
    res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept');
    res.header('Access-Control-Allow-Methods', 'GET, PUT, POST, DELETE, OPTIONS');
    next();
});

// Options
app.options('*', (req, res) => {
    res.sendStatus(200);
});

// Listen
var server = app.listen(process.env.port || 3000, function () {
    console.log("Teams Scraping App is listening to PORT:" + server.address().port);
});


// AD auth param
var param_auth = {
    url: `https://login.microsoftonline.com/${loginUrl}/oauth2/v2.0/authorize`, // e.g. hoge.onmicrosoft.com
    qs: {
        client_id: ADClientId,
        redirect_uri: "http://localhost:3000/token",
        response_type: "code",
        scope: "user.read group.readwrite.all user.readwrite.all user.readwrite",
        state: "1234"
    }
}

// teams param
const param_team = {
    url: "",
    auth: {
        bearer: access_token
    }
}

// reply param
const param_team_reply = {
    url: "",
    auth: {
        bearer: access_token
    }
}

// graph token
var param_token = {
    url: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    headers: {
        "Content-type": "application/x-www-form-urlencoded",
    },
    form: {
        client_id: ADClientId,
        client_secret: ADSecret,
        code: "dummy",
        redirect_uri: "http://localhost:3000/token",
        grant_type: "authorization_code",
        scope: "user.read"
    }
}

class Team {
    constructor() {
        this.teamId = "";
        this.teamName = "";
        this.channels = [];
    }
}

class Channel {
    constructor() {
        this.channelName = "";
        this.channelId = "";
    }
}

/**
 * Scraping Teams Data Endpoint
 */
app.get("/", function (req, res, next) {
    request.get(param_auth, (e, r, b) => {
        res.redirect(r.request.href);
    })
});

/**
 * Scraping Teams Data Endpoint
 */
app.get("/token", function (req, res, next) {

    // Setting Code
    param_token.form.code = req.query.code;

    request.post(param_token, (e, r, b) => {
        // Scraping
        b = JSON.parse(b);
        access_token = b.access_token;
        res.redirect('http://localhost:3000/teams')
    })


});


// Root(Health Check)
app.get("/teams", async function (req, res, next) {

    let teamlist = [];

    // 全チーム取得
    res.send("now scraping...");
    param_team.auth.bearer = access_token;
    param_team_reply.bearer = access_token;
    param_team.url = "https://graph.microsoft.com/beta/groups?$filter=resourceProvisioningOptions/Any(x:x eq 'Team')";
    request.get(param_team, async (e, r, teams) => {
        teams = JSON.parse(teams).value;
        for (team of teams) {
            // 全チャンネル取得
            let newTeam = new Team();
            var teamName = team.displayName;
            var teamId = team.id;
            newTeam.teamId = teamId;
            newTeam.teamName = teamName;
            newTeam.channels = [];

            param_team.url = "https://graph.microsoft.com/beta/teams/" + teamId + "/channels";
            let channels = await syncRequest(param_team);
            channels = JSON.parse(channels);
            if (!channels.error) {
                // 権限があるチーム
                channels = channels.value;

                // 全スレッド
                for (channel of channels) {
                    let newChannel = new Channel();
                    let channelId = channel.id;
                    let channelName = channel.displayName;
                    newChannel.channelId = channelId;
                    newChannel.channelName = channelName;
                    newTeam.channels.push(newChannel);
                    console.log(channelName);
                }
            }
            teamlist.push(newTeam);
        }

        executePromises = [];

        for (team of teamlist) {
            console.log(team.teamName + "is scraping..")
            for (channel of team.channels) {
                // 取得
                await getMessageRecursiveAsync(team.teamId, team.teamName, channel.channelId, channel.channelName);
            }
        }

        allText = "teamId\tteamName\tchannelId\tchannelName\tuserId\tname\tcontent";
        for (msg of messageText) {
            allText += "\n" + msg.teamId;
            allText += "\t" + msg.teamName;
            allText += "\t" + msg.channelId;
            allText += "\t" + msg.channelName;
            allText += "\t" + msg.userId;
            allText += "\t" + msg.name;
            allText += "\t" + msg.content;
        }
        fs.writeFileSync('teams.tsv', allText);
        console.log("done!")

    })
});

/**
 * getMessageRecursiveAsync
 * @param {*} team[]
 */
async function getMessageRecursiveAsync(teamId, teamName, channelId, channelName) {
    return new Promise(async (resolve, reject) => {
        isMessageDone = false;

        // URL Build
        param_team.url = "https://graph.microsoft.com/beta/teams/" +
            teamId + "/channels/" + urlencode(channelId) + "/messages";

        while (!isMessageDone) {
            messages = await syncRequest(param_team);
            if (isJSON(messages)) {
                var nextLink = JSON.parse(messages)["@odata.nextLink"];
                if (nextLink && nextLink != "") {
                    param_team.url = nextLink;
                } else {
                    // 終了
                    isMessageDone = true;
                    console.log(channelName + "has been readed.")
                }
                messages = JSON.parse(messages).value;

                if (messages && messages != undefined) {
                    for (message of messages) {
                        if (message.from.user) {
                            // たまにユーザーがない
                            var id = message.id;
                            var content = message.body.content;
                            content = content.replace(/<("[^"]*"|'[^']*'|[^'">])*>/g, ''); // HTML Tag
                            content = content.replace(/((http:|https:)\/\/[\x21-\x7e]+)/gi, ""); // URL
                            content = content.replace(/&nbsp;/g, ""); // space
                            content = content.replace(/........-....-....-....-............/g, ""); // attachment
                            content = content.replace(/\n/g, ""); // return
                            content = content.replace(/\r/g, ""); // return
                            content = content.replace(/\t/g, ""); // tab
                            var userId = message.from.user.id;
                            var name = message.from.user.displayName;
                            if (content != "") {
                                messageText.push({
                                    teamId: teamId,
                                    teamName: teamName,
                                    channelId: channelId,
                                    channelName: channelName,
                                    content: content,
                                    userId: userId,
                                    name: name
                                });
                            }
                            await getReplyMessageRecursiveAsync(teamId, teamName, channelId, channelName, id);

                        }
                    }
                }
            }

        }

        resolve();
    });
}


/**
 * getMessageRecursiveAsync
 * @param {*} team[]
 */
async function getReplyMessageRecursiveAsync(teamId, teamName, channelId, channelName, msgId) {
    return new Promise(async (resolve, reject) => {
        isMessageDoneReply = false;

        // URL Build
        param_team_reply.url = "https://graph.microsoft.com/beta/teams/" +
            teamId + "/channels/" + urlencode(channelId) + "/messages/" + msgId + "/replies";

        while (!isMessageDoneReply) {
            messages = await syncRequest(param_team_reply);
            if (isJSON(messages)) {
                var nextLink = JSON.parse(messages)["@odata.nextLink"];
                if (nextLink && nextLink != "") {
                    param_team_reply.url = nextLink;
                } else {
                    // 終了
                    isMessageDoneReply = true;
                }
                messages = JSON.parse(messages).value;

                if (messages && messages != undefined) {
                    for (message of messages) {
                        if (message.from.user) {
                            // たまにユーザーがない
                            var content = message.body.content;
                            content = content.replace(/<("[^"]*"|'[^']*'|[^'">])*>/g, ''); // HTML Tag
                            content = content.replace(/((http:|https:)\/\/[\x21-\x7e]+)/gi, ""); // URL
                            content = content.replace(/&nbsp;/g, ""); // space
                            content = content.replace(/........-....-....-....-............/g, ""); // attachment
                            content = content.replace(/\n/g, ""); // return
                            content = content.replace(/\r/g, ""); // return
                            var userId = message.from.user.id;
                            var name = message.from.user.displayName;
                            if (content != "") {
                                messageText.push({
                                    teamId: teamId,
                                    teamName: teamName,
                                    channelId: channelId,
                                    channelName: channelName,
                                    content: content,
                                    userId: userId,
                                    name: name
                                });
                            }

                        }
                    }
                }
            }

        }

        resolve();
    });
}

function syncRequest(param) {
    return new Promise((resolve, reject) => {
        request.get(param, async (e, r, b) => {
            resolve(b)
        })
    })
}


function urlencode(str) {
    return encodeURIComponent(str).replace(/[!'()*]/g, function (c) {
        return '%' + c.charCodeAt(0).toString(16).toUpperCase();
    });
}

function isJSON(arg) {
    arg = (typeof arg === "function") ? arg() : arg;
    if (typeof arg !== "string") {
        return false;
    }
    try {
        arg = (!JSON) ? eval("(" + arg + ")") : JSON.parse(arg);
        return true;
    } catch (e) {
        return false;
    }
};


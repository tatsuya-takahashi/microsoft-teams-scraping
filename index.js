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
var refresh_token = "";
var messageText = [];
var relations = [];
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
    url: "https://login.microsoftonline.com/toyosoft1.onmicrosoft.com/oauth2/v2.0/authorize",
    qs: {
        client_id: "{{:clientId}}",
        redirect_uri: "http://localhost:3000/token",
        response_type: "code",
        scope: "user.read group.readwrite.all user.readwrite.all user.readwrite offline_access",
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
let param_team_reply = {
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
        client_id: "{{:clientId}}",
        client_secret: "{{:clientSecret}}",
        code: "dummy",
        redirect_uri: "http://localhost:3000/token",
        grant_type: "authorization_code",
        scope: "user.read"
    }
}

// graph token
var param_token_refresh = {
    url: "https://login.microsoftonline.com/common/oauth2/v2.0/token",
    headers: {
        "Content-type": "application/x-www-form-urlencoded",
    },
    form: {
        client_id: "{{:clientId}}",
        client_secret: "{{:clientSecret}}",
        refresh_token: "dummy",
        redirect_uri: "http://localhost:3000/token",
        grant_type: "refresh_token",
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
        refresh_token = b.refresh_token;
        res.redirect('http://localhost:3000/teams')
    })


});


// Root(Health Check)
app.get("/teams", async function (req, res, next) {

    let teamlist = [];

    // 全チーム取得
    res.send("now scraping...");
    param_team.auth.bearer = access_token;
    param_team_reply.auth.bearer = access_token;
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

            // 1チーム終わるたびにReflesh
            await refreshToken();
        }

        // alltext
        allText = "teamId\tteamName\tchannelId\tchannelName\tuserId\tname\tcontent";
        for (msg of messageText) {
            allText += "\n" + msg.teamId;
            allText += "\t" + msg.teamName;
            allText += "\t" + msg.channelId;
            allText += "\t" + msg.channelName;
            allText += "\t" + msg.userId;
            allText += "\t" + msg.name;
            allText += "\t" + msg.content;
            allText += "\t" + msg.date;
        }
        fs.writeFileSync('teams.tsv', allText);

        // relation
        allText = "relationType\tfromUserId\ttoUserId";
        for (rel of relations) {
            allText += "\n" + rel.relationType;
            allText += "\t" + rel.fromUserId;
            allText += "\t" + rel.toUserId;
        }
        fs.writeFileSync('teams_relations.tsv', allText);
        console.log("done!")

    })
});

function refreshToken() {
    return new Promise(async (resolve, reject) => {
        // Setting Code
        param_token_refresh.form.refresh_token = refresh_token;

        // Token更新
        b = await syncRequestPost(param_token_refresh);
        b = JSON.parse(b);
        access_token = b.access_token;
        refresh_token = b.refresh_token;
        param_team.auth.bearer = access_token;
        param_team_reply.auth.bearer = access_token;
        resolve();
    })
}

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
                    console.log(JSON.parse(messages));
                    console.log(channelName + "has been readed.")
                }
                messages = JSON.parse(messages).value;

                if (messages && messages != undefined) {
                    for (message of messages) {
                        if (message.from && message.from.user) {
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
                            var date = message.createdDateTime.toString();
                            if (content != "") {
                                messageText.push({
                                    teamId: teamId,
                                    teamName: teamName,
                                    channelId: channelId,
                                    channelName: channelName,
                                    content: content,
                                    userId: userId,
                                    name: name,
                                    date: date
                                });
                            }

                            // likeの取得
                            for (likeUser of message.reactions) {
                                if (userId != likeUser.user.user.id) {
                                    relations.push({
                                        relationType: likeUser.reactionType,
                                        fromUserId: likeUser.user.user.id,
                                        toUserId: userId
                                    });
                                }
                            }

                            await getReplyMessageRecursiveAsync(teamId, teamName, channelId, channelName, id, userId);

                        } else {
                            // console.log(message);
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
async function getReplyMessageRecursiveAsync(teamId, teamName, channelId, channelName, msgId, replyToUserId) {
    return new Promise(async (resolve, reject) => {
        isMessageDoneReply = false;

        // // URL Build
        // url = "https://graph.microsoft.com/beta/teams/" + teamId + "/channels/" + urlencode(channelId) + "/messages/" + msgId + "/replies";
        // param_team_reply = {
        //     url: url,
        //     auth: {
        //         bearer: token
        //     }
        // }
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
                        if (message.from && message.from.user) {
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
                            var date = message.createdDateTime.toString();
                            if (content != "") {
                                messageText.push({
                                    teamId: teamId,
                                    teamName: teamName,
                                    channelId: channelId,
                                    channelName: channelName,
                                    content: content,
                                    userId: userId,
                                    name: name,
                                    date: date
                                });
                            }

                            // reply
                            if (userId != replyToUserId) {
                                relations.push({
                                    relationType: 'reply',
                                    fromUserId: userId,
                                    toUserId: replyToUserId
                                });
                            }

                            // likeの取得
                            for (likeUser of message.reactions) {
                                if (userId != likeUser.user.user.id) {
                                    relations.push({
                                        relationType: likeUser.reactionType,
                                        fromUserId: likeUser.user.user.id,
                                        toUserId: userId
                                    });
                                }
                            }

                        }
                    }
                }
            } else {
                isMessageDoneReply = true;
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

function syncRequestPost(param) {
    return new Promise((resolve, reject) => {
        request.post(param, async (e, r, b) => {
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


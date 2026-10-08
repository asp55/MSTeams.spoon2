// MSTeams/init.js

//-----------------------------------------
// Logger
//-----------------------------------------
const Logger = {
    log:(...params)=>{
        console.log("[MSTeams]", ...params);
    },
    debug:(...params)=>{
        console.debug("[MSTeams]", ...params);
    },
    error:(...params)=>{
        console.error("[MSTeams]", ...params);
    },
    info:(...params)=>{
        console.info("[MSTeams]", ...params);
    },
    warn:(...params)=>{
        console.warn("[MSTeams]", ...params);
    },

} 

//-----------------------------------------
// End Logger
//-----------------------------------------


//-----------------------------------------
// Declare Variables
//-----------------------------------------


// MSTeams.permissions
// Read-only object of meetingPermissions provided by Teams or null if not connected to teams.
let meetingPermissions = null;

// MSTeams.state
// Read-only object of meetingState properties provided by Teams or null if not connected to teams.
let meetingState = null;

let updateCallback = ()=>{};

// private variable to track if spoon is already running or not. (Makes it easier to find local variables)
let running = false;

//-----------------------------------------
// End of Declare Variables
//-----------------------------------------


//-----------------------------------------
// Teams Monitor
//-----------------------------------------

let teamsPairing = false;
let teamsWebsocket = null;
let teamsConnectionId = 0;


function disconnectFromTeams() {
    if(teamsWebsocket) {
        teamsWebsocket.close();
        teamsWebsocket = null;
    }
}


let requestID = 0;
function sendRequest(msg) {
    requestID++;
    const requestMSG = JSON.stringify({
        requestId: requestID,
        action: msg.action ?? "",
        parameters: msg.parameters ?? {}
    })
    Logger.debug("sendRequest", requestMSG);
    if(teamsWebsocket) teamsWebsocket.send(requestMSG);
}

let closedCount = 0;

function onTeamsClose(code, reason) {
    Logger.log(`Teams WebSocket Closed. {code:${code}, reason:${reason}`);
    teamsWebsocket = null;
    teamsPairing = false;
    if(running) {
        closedCount++;
        if(closedCount > 3) {
            Logger.warn("Teams WebSocket closed multiple times in a row")
            Logger.warn("This likely means this app was blocked from the Third-party app API in teams.")
            Logger.warn("Go to Settings > Privacy > Third-party app API > Manage API and remove the application from block.")
            Logger.warn("Then restart this spoon.")
        }
        else {
            Logger.info("Teams not available, retrying in 5 seconds");
            hs.timer.doAfter(5, connectToTeams);
        }
    }

}
function onTeamsOpen() {
    closedCount = 0;
    Logger.debug(`Connected to Teams local API`);
}

function onTeamsMessage(message) {
    Logger.debug(`Teams WebSocket Received`, message);

    try {
        const parsed = JSON.parse(message);

        if(parsed.tokenRefresh) {
            Logger.debug("Teams token refreshed");
            hs.userdefaults.set("MSTeams.teamsToken", parsed.tokenRefresh)
        }

        if(parsed.meetingUpdate) {
            if(parsed.meetingUpdate.meetingPermissions) {
                meetingPermissions = parsed.meetingUpdate.meetingPermissions;
                Logger.debug("Got new meeting permissions", meetingPermissions);

                if(parsed.meetingUpdate.meetingPermissions.canPair && !teamsPairing) {
                    Logger.debug("Sending pairing request");
                    teamsPairing = true;
                    MSTeams.pair();
                }
            }

            if(parsed.meetingUpdate.meetingState) {
                meetingState = parsed.meetingUpdate.meetingState;
                Logger.debug("Got new meeting state", meetingState);
            }

            updateCallback();
        }

        if(parsed.response && parsed.response === "Pairing response resulted in no action") {
            Logger.debug("Didn't pair. Will try again next meeting.");
            teamsPairing = false
        }
    }
    catch(e) {
        Logger.warn("Failed to parse Teams message: ", message);
    }



}
function onTeamsError(err) {
    Logger.log(`Teams WebSocket Error`, err);

    teamsWebsocket = null;
    if(running) {
        Logger.debug("Teams not available, retrying in 30 seconds")
        hs.timer.doAfter(30, connectToTeams);
    }
}

function connectToTeams() {
    Logger.log(`connectToTeams()`);

    // Increment the connection ID before closing, so any callbacks from the
    // previous connection are ignored even if close() fires synchronously.
    teamsConnectionId++;
    const myId = teamsConnectionId;
    if(teamsWebsocket) {
        teamsWebsocket.close();
        teamsWebsocket = null;
    }

    const token = hs.userdefaults.get("MSTeams.teamsToken") ?? "";
    const manufacturer = "Hammerspoon2";
    const device = "MSTeams.spoon2";
    const app = "MSTeams.spoon2";
    const url = `ws://localhost:8124?token=${token}&protocol-version=2.0.0&manufacturer=${manufacturer}&device=${device}&app=${app}&app-version=${MSTeams.version}`;

    teamsWebsocket = hs.http.openWebSocket(url)
        .setOpenCallback(onTeamsOpen)
        .setMessageCallback(onTeamsMessage)
        .setCloseCallback(onTeamsClose)
        .setErrorCallback(onTeamsError);


}


//-----------------------------------------
// End Teams Monitor
//-----------------------------------------




const MSTeams = {
    get permissions() {
        return meetingPermissions;
    },
    get state() {
        return meetingState;
    },
};

MSTeams.init = function () {
    // Called automatically by hs.loadSpoon(). author/description/version are already set on
    // `this` by this point - safe to log or read, but this could run more than once, so avoid
    // anything that shouldn't happen twice.
    Logger.log(`v${this.version} initialized`)
};

MSTeams.start = function () {
    Logger.log("start()");
    if(!running) {
        running = true;
        if(!teamsWebsocket) connectToTeams();
    }
    else {
        Logger.warn(`Already Started`);
    }

    return module.exports;
}

MSTeams.stop = function() {
    Logger.log("stop()");

    running = false;
    closedCount = 0;
    meetingPermissions = null;
    meetingState = null;

    disconnectFromTeams();

    return module.exports;
};

MSTeams.onUpdate = function (callback) {
    if(typeof callback === 'function') {
        updateCallback = callback;
    }
    else {
        throw new Error('callback must be a function')
    }

    return module.exports;
};




//-----------------------------------------
// Teams Command Methods
//-----------------------------------------

MSTeams.queryState = function () {
    if(!running) {
        Logger.warn("spoon MSTeams must be running to perform actions. Run MSTeams.start()");
    }
    else if(!teamsWebsocket) {
        Logger.warn("Not connected to teams");
    }
    else {
        sendRequest({action:"query-state", parameters:{}});
    }
    
    return module.exports;
}

const permission = {
   canLeave: "canLeave",
   canPair: "canPair",
   canReact: "canReact",
   canStopSharing: "canStopSharing",
   canToggleBlur: "canToggleBlur",
   canToggleChat: "canToggleChat",
   canToggleHand: "canToggleHand",
   canToggleMute: "canToggleMute",
   canToggleShareTray: "canToggleShareTray",
   canToggleVideo: "canToggleVideo"
}

function canAct(requiredPermission) {
    if(!running) {
        Logger.warn("spoon MSTeams must be running to perform actions. Run MSTeams.start()");
        return false;
    }

    if(!teamsWebsocket) {
        Logger.warn("Not connected to teams");
        return false;
    }

    if(!meetingPermissions) {
        Logger.warn("Not currently in a teams meeting");
        return false;
    }

    if(requiredPermission) {
        if(!permission[requiredPermission]) {
            Logger.error("Invalid permission");
            return false;
        }
        else if(!meetingPermissions[requiredPermission]){
            Logger.warn("Don't have necessary permission in current meeting.", `${requiredPermission} = ${meetingPermissions[requiredPermission]}`);
            return false;
        }
    }

    return true;
}

MSTeams.pair = function () {
    if(canAct(permission.canPair)) {
        sendRequest({action:"pair", parameters:{}})
    }
    
    return module.exports;
}

MSTeams.toggleMute = function () {
    if(canAct(permission.canToggleMute)) {
        sendRequest({action:"toggle-mute", parameters:{}})
    }
    
    return module.exports;
}

MSTeams.mute = function () {
    if(canAct(permission.canToggleMute)) {
        sendRequest({action:"mute", parameters:{}})
    }
    
    return module.exports;
}

MSTeams.unmute = function () {
    if(canAct(permission.canToggleMute)) {
        sendRequest({action:"unmute", parameters:{}})
    }
    
    return module.exports;
}

MSTeams.toggleVideo = function () {
    if(canAct(permission.canToggleVideo)) {
        sendRequest({action:"toggle-video", parameters:{}})
    }
    
    return module.exports;
}

MSTeams.showVideo = function () {
    if(canAct(permission.canToggleVideo)) {
        sendRequest({action:"show-video", parameters:{}})
    }
    
    return module.exports;
}

MSTeams.hideVideo = function () {
    if(canAct(permission.canToggleVideo)) {
        sendRequest({action:"hide-video", parameters:{}})
    }
    
    return module.exports;
}

MSTeams.stopSharing = function () {
    if(canAct(permission.canStopSharing)) {
        sendRequest({action:"stop-sharing", parameters:{}})
    }
    
    return module.exports;
}

MSTeams.toggleBlurBackground = function () {
    if(canAct(permission.canToggleBlur)) {
        sendRequest({action:"toggle-background-blur", parameters:{}})
    }
    
    return module.exports;
}

MSTeams.blurBackground = function () {
    if(canAct(permission.canToggleBlur)) {
        sendRequest({action:"blur-background", parameters:{}})
    }
    
    return module.exports;
}

MSTeams.unblurBackground = function () {
    if(canAct(permission.canToggleBlur)) {
        sendRequest({action:"unblur-background", parameters:{}})
    }
    
    return module.exports;
}

MSTeams.toggleHand = function () {
    if(canAct(permission.canToggleHand)) {
        sendRequest({action:"toggle-hand", parameters:{}})
    }
    
    return module.exports;
}

MSTeams.raiseHand = function () {
    if(canAct(permission.canToggleHand)) {
        sendRequest({action:"raise-hand", parameters:{}})
    }
    
    return module.exports;
}

MSTeams.lowerHand = function () {
    if(canAct(permission.canToggleHand)) {
        sendRequest({action:"lower-hand", parameters:{}})
    }
    
    return module.exports;
}

MSTeams.reactLike = function (reaction) {
    if(canAct(permission.canReact)) {
        sendRequest({action:"send-reaction", parameters:{type:"like"}})
    }
    
    return module.exports;
}

MSTeams.reactLove = function (reaction) {
    if(canAct(permission.canReact)) {
        sendRequest({action:"send-reaction", parameters:{type:"love"}})
    }
    
    return module.exports;
}

MSTeams.reactApplause = function (reaction) {
    if(canAct(permission.canReact)) {
        sendRequest({action:"send-reaction", parameters:{type:"applause"}})
    }
    
    return module.exports;
}

MSTeams.reactLaugh = function (reaction) {
    if(canAct(permission.canReact)) {
        sendRequest({action:"send-reaction", parameters:{type:"laugh"}})
    }
    
    return module.exports;
}

MSTeams.reactWow = function (reaction) {
    if(canAct(permission.canReact)) {
        sendRequest({action:"send-reaction", parameters:{type:"wow"}})
    }
    
    return module.exports;
}

MSTeams.leaveCall = function () {
    if(canAct(permission.canLeave)) {
        sendRequest({action:"leave-call", parameters:{}})
    }
    
    return module.exports;
}


MSTeams.toggleChat = function () {
    if(canAct(permission.canToggleChat)) {
        sendRequest({action:"toggle-ui", parameters:{type:"chat"}})
    }
    
    return module.exports;
}

MSTeams.toggleShareTray = function () {
    if(canAct(permission.canToggleShareTray)) {
        sendRequest({action:"toggle-ui", parameters:{type:"share-tray"}})
    }
    
    return module.exports;
}

MSTeams.customRequest = function (action, parameters) {
    if(canAct()) {
        sendRequest({action:action, parameters:parameters ?? {}})
    }
}

//-----------------------------------------
// End Teams Command Methods
//-----------------------------------------


module.exports = MSTeams;
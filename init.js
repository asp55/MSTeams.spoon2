// MSTeams/init.js

function init() {
    // Called automatically by hs.loadSpoon(). author/description/version are already set on
    // `this` by this point - safe to log or read, but this could run more than once, so avoid
    // anything that shouldn't happen twice.
    console.log(`[MSTeams] v${this.version} initialized`)
}

function start() {
    console.log("[MSTeams] started")
}

function stop() {
    console.log("[MSTeams] stopped")
}


module.exports = { init, start, stop }
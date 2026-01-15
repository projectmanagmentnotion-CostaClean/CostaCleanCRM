function __pingLogs() {
  console.log("PING console.log => " + new Date().toISOString());
  Logger.log("PING Logger.log => " + new Date());
}

import fs from "fs";
import path from "path";

class FileLogger {
  constructor() {
    this.logFile = path.resolve("./mcp_sharepoint.log");
  }

  formatMessage(level, message, ...args) {
    const timestamp = new Date().toISOString();
    const formattedMessage =
      args.length > 0
        ? `${message} ${args
            .map((arg) =>
              typeof arg === "object" ? JSON.stringify(arg) : String(arg)
            )
            .join(" ")}`
        : message;
    return `${timestamp} - mcp_sharepoint - ${level} - ${formattedMessage}`;
  }

  writeLog(level, message, ...args) {
    const logMessage = this.formatMessage(level, message, ...args);

    // Write to console
    console.log(logMessage);

    // Write to file
    try {
      fs.appendFileSync(this.logFile, logMessage + "\n");
    } catch (error) {
      console.error("Failed to write to log file:", error);
    }
  }

  info(message, ...args) {
    this.writeLog("INFO", message, ...args);
  }

  error(message, ...args) {
    this.writeLog("ERROR", message, ...args);
  }

  warn(message, ...args) {
    this.writeLog("WARN", message, ...args);
  }

  debug(message, ...args) {
    this.writeLog("DEBUG", message, ...args);
  }
}

export const logger = new FileLogger();

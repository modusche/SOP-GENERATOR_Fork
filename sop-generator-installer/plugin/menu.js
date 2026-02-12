module.exports = function (electronApp, menuState) {
  var spawn = require("child_process").spawn;
  var net = require("net");
  var path = require("path");
  var fs = require("fs");
  var os = require("os");
  var { app } = require("electron");

  var SOP_SERVER_PORT = 8000;
  var serverProcess = null;

  // Development: plugin backend path. Production: bundled exe in resources.
  var isDev = !app.isPackaged;
  var backendDir;
  var sopServerExe;

  if (isDev) {
    backendDir = path.join(__dirname, "..", "backend");
    sopServerExe = path.join(backendDir, "sop_server.py");
  } else {
    sopServerExe = path.join(
      process.resourcesPath,
      "sop-server",
      process.platform === "win32" ? "sop-server.exe" : "sop-server"
    );
  }

  function isPortInUse(port, callback) {
    var server = net.createServer();
    server.once("error", function () {
      callback(true);
    });
    server.once("listening", function () {
      server.close(function () {
        callback(false);
      });
    });
    server.listen(port, "127.0.0.1");
  }

  isPortInUse(SOP_SERVER_PORT, function (inUse) {
    if (inUse) {
      console.log(
        "[SOP Plugin] Port " + SOP_SERVER_PORT + " already in use - backend already running"
      );
      return;
    }

    if (!fs.existsSync(sopServerExe)) {
      console.error("[SOP Plugin] Backend not found at: " + sopServerExe);
      if (isDev) {
        console.error("[SOP Plugin] Development: ensure Python is installed and backend files exist.");
      } else {
        console.error("[SOP Plugin] Production: sop-server executable not found in resources.");
      }
      return;
    }

    console.log("[SOP Plugin] Starting SOP Generator backend...");

    if (isDev) {
      var pythonCmd = process.platform === "win32" ? "python" : "python3";
      serverProcess = spawn(
        pythonCmd,
        [sopServerExe, String(SOP_SERVER_PORT)],
        {
          stdio: "ignore",
          detached: false,
          windowsHide: true,
          cwd: path.dirname(sopServerExe),
        }
      );
    } else {
      serverProcess = spawn(sopServerExe, [String(SOP_SERVER_PORT)], {
        stdio: "ignore",
        detached: false,
        windowsHide: true,
        cwd: path.dirname(sopServerExe),
      });
    }

    serverProcess.on("error", function (err) {
      console.error("[SOP Plugin] Failed to start backend:", err.message);
      serverProcess = null;
    });

    serverProcess.on("exit", function (code) {
      console.log("[SOP Plugin] Backend exited with code", code);
      serverProcess = null;
    });
  });

  electronApp.on("quit", function () {
    if (serverProcess) {
      console.log("[SOP Plugin] Stopping SOP Generator backend...");
      serverProcess.kill();
      serverProcess = null;
    }
  });

  return [
    {
      label: "Generate SOP Document",
      accelerator: "CommandOrControl+Shift+G",
      enabled: function () {
        return menuState.bpmn;
      },
      action: function () {
        electronApp.emit("menu:action", "generate-sop");
      },
    },
  ];
};

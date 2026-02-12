# SOP Generator - Camunda Modeler Plugin

A Camunda Modeler plugin that generates formatted Microsoft Word SOP (Standard Operating Procedure) documents directly from BPMN diagrams. The plugin is **bundled with Camunda Modeler** — no manual installation required. Simply open a BPMN diagram, click the **EXPORT** button in the status bar, review the auto-populated metadata in a modal form, and download a properly formatted `.docx` file following Guideline V2 specifications.

---

## System Requirements

- **Windows 10** or later (or Linux/Mac)
- **Camunda Modeler** with SOP Generator plugin bundled
- **No Python installation required** — backend is bundled as standalone executable

---

## Installation

**No installation needed!** The SOP Generator plugin is bundled with Camunda Modeler. The backend server starts automatically when Camunda Modeler launches.

---

## Usage

1. Open a BPMN diagram in Camunda Modeler
2. Click the **EXPORT** button in the status bar (visible when BPMN tab is active)
   - Alternatively, press **Ctrl+Shift+G** (keyboard shortcut still works)
3. A modal form appears with auto-populated metadata extracted from the diagram:
   - Process name, code, purpose, scope
   - Abbreviations and definitions
   - Referenced documents (auto-generated from lane names)
   - General policies
4. Edit any fields as needed
5. Click **Generate & Download .docx**
6. The Word document is downloaded and the modal closes

---

## Uninstallation

The plugin is part of Camunda Modeler. To disable it, remove the plugin directory from the Camunda Modeler installation.

Generated documents and history data in `%LOCALAPPDATA%\SOP_Generator\` are preserved.

---

## How the Camunda Modeler Plugin System Works

This section documents the plugin architecture for anyone who wants to create additional Camunda Modeler plugins.

### Plugin Discovery

Camunda Modeler scans the following directory at startup:

```
%APPDATA%\camunda-modeler\resources\plugins\
```

Each subdirectory containing an `index.js` file is loaded as a plugin. No registration or configuration file is needed — just place a folder with `index.js` in the plugins directory.

### Plugin Entry Point: `index.js`

Every plugin must have an `index.js` that exports a module with up to three properties:

```javascript
module.exports = {
  name: 'My Plugin',       // Display name (shown in logs)
  script: './client.js',   // Client-side script (runs in browser/renderer context)
  menu: './menu.js'        // Menu module (runs in Node.js/Electron main process)
};
```

All three properties are optional — you can have a menu-only plugin, a client-only plugin, or both.

### Menu Module: `menu.js`

The menu module runs in **Electron's main process** with full Node.js access. It receives two parameters:

```javascript
module.exports = function(electronApp, menuState) {
  // Full Node.js access: require('fs'), require('child_process'), etc.

  return [
    {
      label: 'My Action',                      // Menu item text
      accelerator: 'CommandOrControl+Shift+M',  // Keyboard shortcut
      enabled: function() {
        return menuState.bpmn;  // Only enable when a BPMN diagram is open
      },
      action: function() {
        // Emit an action that the client module can handle
        electronApp.emit('menu:action', 'my-custom-action');
      }
    }
  ];
};
```

**Key points:**
- Runs in Node.js — full access to `fs`, `child_process`, `net`, `path`, `os`, etc.
- `menuState.bpmn` is `true` when a BPMN diagram tab is active
- `electronApp.emit('menu:action', 'action-name')` bridges to the client module
- Code here executes once at startup (outside the return), making it a good place to spawn background processes

### Client Module: `client.js`

The client module runs in the **browser/renderer context** with DOM access. It registers a bpmn-js module:

```javascript
(function() {
  'use strict';

  // Define a bpmn-js module using dependency injection
  function MyModule(bpmnjs, editorActions) {
    // Register a handler for the action emitted by menu.js
    editorActions.register('my-custom-action', function() {

      // Access the current diagram
      bpmnjs.saveXML({ format: true }).then(function(result) {
        var xml = result.xml;
        // Do something with the BPMN XML...
        console.log('Got XML:', xml.substring(0, 100));
      });

    });
  }

  // Declare dependencies to inject (must match parameter names)
  MyModule.$inject = ['bpmnjs', 'editorActions'];

  // Package as a bpmn-js module
  var myModule = {
    __init__: ['myModule'],           // Key must match the property name below
    myModule: ['type', MyModule]      // Registers MyModule as the 'myModule' service
  };

  // Register with Camunda Modeler's plugin system
  var plugins = window.plugins || [];
  window.plugins = plugins;
  plugins.push({
    plugin: myModule,
    type: 'bpmn.modeler.additionalModules'
  });

})();
```

**Key points:**
- Runs in Chromium — DOM access, `fetch`, `window`, `document`, etc.
- No Node.js APIs (`fs`, `child_process` are not available)
- Uses an IIFE to avoid polluting the global scope
- The `$inject` array declares which bpmn-js services to inject

### Communication: Menu to Client

The bridge between the Electron main process (menu.js) and the renderer (client.js):

```
menu.js                                  client.js
────────                                 ─────────
electronApp.emit('menu:action',    →     editorActions.register(
  'my-custom-action')                      'my-custom-action', handler)
```

1. Menu defines action: `electronApp.emit('menu:action', 'action-name')`
2. Camunda Modeler internally routes it to the active editor tab
3. Client handles it: `editorActions.register('action-name', function() { ... })`

### Available Dependency Injection Targets

These services can be injected into your client module:

| Service            | Description                                       |
|--------------------|---------------------------------------------------|
| `bpmnjs`           | The bpmn-js modeler instance                      |
| `editorActions`    | Register/trigger custom actions                   |
| `eventBus`         | Subscribe to diagram events                       |
| `modeling`         | Programmatically modify diagram elements          |
| `elementRegistry`  | Query diagram elements by ID                      |
| `canvas`           | Access the drawing canvas (zoom, scroll, layers)  |
| `overlays`         | Add HTML overlays on diagram elements             |
| `selection`        | Get/set selected elements                         |

### Available Plugin Types

```javascript
plugins.push({
  plugin: myModule,
  type: 'bpmn.modeler.additionalModules'   // For BPMN modeler
  // Other types:
  // 'dmn.modeler.additionalModules'        // For DMN modeler
  // 'cmmn.modeler.additionalModules'       // For CMMN modeler
});
```

---

## Creating a New Plugin: Step-by-Step

### Example 1: Minimal Menu-Only Plugin

A plugin that just adds a menu entry:

```
%APPDATA%\camunda-modeler\resources\plugins\my-plugin\
├── index.js
└── menu.js
```

**index.js:**
```javascript
module.exports = {
  name: 'My Plugin',
  menu: './menu.js'
};
```

**menu.js:**
```javascript
module.exports = function(electronApp, menuState) {
  return [
    {
      label: 'Say Hello',
      accelerator: 'CommandOrControl+Shift+H',
      enabled: function() { return true; },
      action: function() {
        require('electron').dialog.showMessageBox({
          type: 'info',
          title: 'Hello',
          message: 'Hello from my plugin!'
        });
      }
    }
  ];
};
```

### Example 2: Client Plugin with Diagram Access

A plugin that reads the current diagram:

```
%APPDATA%\camunda-modeler\resources\plugins\diagram-info\
├── index.js
├── menu.js
└── client.js
```

**index.js:**
```javascript
module.exports = {
  name: 'Diagram Info',
  script: './client.js',
  menu: './menu.js'
};
```

**menu.js:**
```javascript
module.exports = function(electronApp, menuState) {
  return [{
    label: 'Show Diagram Info',
    accelerator: 'CommandOrControl+Shift+I',
    enabled: function() { return menuState.bpmn; },
    action: function() { electronApp.emit('menu:action', 'show-diagram-info'); }
  }];
};
```

**client.js:**
```javascript
(function() {
  function DiagramInfo(bpmnjs, editorActions, elementRegistry) {
    editorActions.register('show-diagram-info', function() {
      var elements = elementRegistry.getAll();
      var tasks = elements.filter(function(e) { return e.type === 'bpmn:Task'; });
      var gateways = elements.filter(function(e) { return e.type.indexOf('Gateway') !== -1; });

      window.alert(
        'Diagram has ' + elements.length + ' elements:\n' +
        '  Tasks: ' + tasks.length + '\n' +
        '  Gateways: ' + gateways.length
      );
    });
  }

  DiagramInfo.$inject = ['bpmnjs', 'editorActions', 'elementRegistry'];

  window.plugins = window.plugins || [];
  window.plugins.push({
    plugin: {
      __init__: ['diagramInfo'],
      diagramInfo: ['type', DiagramInfo]
    },
    type: 'bpmn.modeler.additionalModules'
  });
})();
```

### Example 3: Plugin with External Backend

The pattern used by the SOP Generator — a menu module spawns a server process, and the client communicates with it via HTTP:

**menu.js** (in main process — has Node.js access):
```javascript
module.exports = function(electronApp, menuState) {
  var spawn = require('child_process').spawn;
  var path = require('path');

  // Spawn a backend server at startup
  var serverProcess = spawn('python', ['my_server.py'], {
    cwd: path.join(process.env.LOCALAPPDATA, 'MyPlugin', 'backend'),
    stdio: 'ignore',
    windowsHide: true
  });

  // Clean up on exit
  electronApp.on('quit', function() {
    if (serverProcess) serverProcess.kill();
  });

  return [{
    label: 'My Backend Action',
    enabled: function() { return menuState.bpmn; },
    action: function() { electronApp.emit('menu:action', 'my-backend-action'); }
  }];
};
```

**client.js** (in renderer — uses fetch to talk to backend):
```javascript
(function() {
  function MyBackendPlugin(bpmnjs, editorActions) {
    editorActions.register('my-backend-action', function() {
      bpmnjs.saveXML({ format: true }).then(function(result) {
        return fetch('http://localhost:9000/api/process', {
          method: 'POST',
          headers: { 'Content-Type': 'application/json' },
          body: JSON.stringify({ xml: result.xml })
        });
      }).then(function(response) {
        return response.json();
      }).then(function(data) {
        window.alert('Result: ' + data.message);
      });
    });
  }

  MyBackendPlugin.$inject = ['bpmnjs', 'editorActions'];

  window.plugins = window.plugins || [];
  window.plugins.push({
    plugin: {
      __init__: ['myBackendPlugin'],
      myBackendPlugin: ['type', MyBackendPlugin]
    },
    type: 'bpmn.modeler.additionalModules'
  });
})();
```

**Important note on iframes and cross-origin:** If your plugin opens an iframe pointing to `localhost`, the iframe and the parent window are on different origins. Use `postMessage` to communicate between them — the iframe cannot directly access `window.parent.document`.

---

## Architecture

### Directory Layout After Installation

```
%LOCALAPPDATA%\SOP_Generator\
  backend\                         Python backend
    app.py                           Flask application + Word doc generation
    bpmn_parser.py                   BPMN XML parser (Guideline V2 rules)
    sop_server.py                    Headless server entry point (waitress)
    history_manager.py               Generation history storage
    archive_manager.py               Document archive management
    create_template.py               Word template generator
    requirements.txt                 Python dependencies
    final_master_template_2.docx     Word document template
    logo.ico, logo.png               Application icons
    templates/
      index.html                     Standalone web UI
      preview.html                   Preview form (used in Camunda modal)
  history/                         Created at runtime
  archives/                        Created at runtime
  debug.log                        Created at runtime

%APPDATA%\camunda-modeler\resources\plugins\sop-generator\
  index.js                         Plugin entry point
  menu.js                          Spawns Python backend, adds menu entry
  client.js                        bpmn-js module, modal overlay, download bridge
```

### Request Flow

```
User presses Ctrl+Shift+G
  │
  ▼
menu.js emits 'generate-sop' action
  │
  ▼
client.js: editorActions handler fires
  │
  ├─ bpmnjs.saveXML() ──── gets current diagram XML
  │
  ├─ POST /api/upload-xml ──── sends XML to Python backend
  │                              returns session_id + metadata
  │
  ├─ Opens modal with iframe ──── /preview/<session_id>
  │                                 shows pre-populated metadata form
  │
  ├─ User edits metadata, clicks Generate
  │
  ├─ iframe sends form data via postMessage to parent
  │
  ├─ client.js: fetch POST /api/generate-and-download/<session_id>
  │               sends form data to backend
  │               receives .docx binary
  │
  ├─ Creates blob URL, triggers download via <a> element
  │
  └─ Closes modal
```

### Backend Processing

```
BPMN XML ──► bpmn_parser.py
               │
               ├─ Extracts tasks, gateways, flows, lanes, events
               ├─ Applies Guideline V2 formatting rules
               ├─ Generates multi-paragraph step structures
               └─ Returns context dict
                    │
                    ▼
              app.py: create_word_doc_from_template()
               │
               ├─ Renders metadata via docxtpl (Jinja2 for Word)
               ├─ Populates process description table via python-docx
               ├─ Applies formatting (fonts, colors, shading)
               └─ Returns .docx as BytesIO stream
```

---

## Troubleshooting

### "Backend server not available" error

1. Verify Python is in PATH:
   ```
   python --version
   ```
2. Reinstall dependencies:
   ```
   python -m pip install -r "%LOCALAPPDATA%\SOP_Generator\backend\requirements.txt"
   ```
3. Test the backend manually:
   ```
   python "%LOCALAPPDATA%\SOP_Generator\backend\sop_server.py"
   ```
   Then open `http://localhost:8000` in a browser — if it loads, the backend works.
4. Check if port 8000 is in use by another application:
   ```
   netstat -an | findstr 8000
   ```

### Plugin does not appear in Camunda Modeler

1. Verify files exist:
   ```
   dir "%APPDATA%\camunda-modeler\resources\plugins\sop-generator"
   ```
   Should show `index.js`, `menu.js`, `client.js`.
2. Restart Camunda Modeler completely (close all windows)
3. Check logs: **Help** > **Toggle Developer Tools** > **Console** tab — look for `[SOP Plugin]` messages

### Word document formatting issues

1. Ensure the template exists:
   ```
   dir "%LOCALAPPDATA%\SOP_Generator\backend\final_master_template_2.docx"
   ```
2. Regenerate the template if missing:
   ```
   cd "%LOCALAPPDATA%\SOP_Generator\backend"
   python create_template.py
   ```
3. The "Avenir LT Std 45 Book" font is recommended but documents will generate without it

### Port 8000 conflict

The backend uses port 8000. If another application uses this port, the plugin detects it and assumes the backend is already running.

To change the port, edit both files:
- `menu.js` — change `SOP_SERVER_PORT` value
- `client.js` — change `SOP_BACKEND` URL

### Debugging

Open Camunda Modeler's developer console (**Help** > **Toggle Developer Tools**) to see:
- `[SOP Plugin]` — messages from menu.js (backend startup/shutdown)
- `[SOP Generator]` — messages from client.js (XML upload, download, errors)

The Python backend logs to: `%LOCALAPPDATA%\SOP_Generator\debug.log`

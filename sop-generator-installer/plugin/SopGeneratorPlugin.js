(function() {
  'use strict';

  // Get React and Fill from global scope (provided by Camunda Modeler)
  // These may not be available immediately, so we'll access them when needed

  const SOP_BACKEND = 'http://localhost:8000';
  const MAX_RETRIES = 3;
  const RETRY_DELAY_MS = 2000;

  // Store bpmnjs reference for React component to use
  let bpmnjsInstance = null;
  let generateSopFunction = null;

  function postWithRetry(url, body, attempt) {
    return fetch(url, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    }).catch(function(err) {
      if (attempt < MAX_RETRIES) {
        console.log('[SOP Generator] Backend not ready, retrying... (' + (attempt + 1) + '/' + MAX_RETRIES + ')');
        return new Promise(function(resolve) {
          setTimeout(function() {
            resolve(postWithRetry(url, body, attempt + 1));
          }, RETRY_DELAY_MS);
        });
      }
      throw err;
    });
  }

  // ============================================================
  // Modal overlay - renders preview form inside Camunda Modeler
  // ============================================================

  var modalOverlay = null;

  function openModal(previewUrl) {
    closeModal();

    modalOverlay = document.createElement('div');
    modalOverlay.id = 'sop-modal-overlay';
    modalOverlay.style.cssText =
      'position:fixed;top:0;left:0;width:100%;height:100%;' +
      'background:rgba(0,0,0,0.5);z-index:99999;display:flex;' +
      'align-items:center;justify-content:center;';

    var modalBox = document.createElement('div');
    modalBox.style.cssText =
      'background:#fff;border-radius:10px;width:90%;max-width:950px;height:85%;' +
      'display:flex;flex-direction:column;box-shadow:0 8px 40px rgba(0,0,0,0.3);overflow:hidden;';

    // Title bar
    var titleBar = document.createElement('div');
    titleBar.style.cssText =
      'display:flex;align-items:center;justify-content:space-between;' +
      'padding:10px 16px;background:#343a40;color:#fff;flex-shrink:0;';

    var title = document.createElement('span');
    title.textContent = 'SOP Generator';
    title.style.cssText = 'font-weight:600;font-size:14px;';

    var closeBtn = document.createElement('button');
    closeBtn.textContent = '\u2715';
    closeBtn.style.cssText =
      'background:none;border:none;color:#fff;font-size:18px;cursor:pointer;' +
      'padding:0 4px;line-height:1;';
    closeBtn.onclick = closeModal;

    titleBar.appendChild(title);
    titleBar.appendChild(closeBtn);

    // Iframe
    var iframe = document.createElement('iframe');
    iframe.src = previewUrl;
    iframe.style.cssText = 'flex:1;border:none;width:100%;';

    modalBox.appendChild(titleBar);
    modalBox.appendChild(iframe);
    modalOverlay.appendChild(modalBox);
    document.body.appendChild(modalOverlay);

    // Close on backdrop click
    modalOverlay.addEventListener('click', function(e) {
      if (e.target === modalOverlay) closeModal();
    });

    // Close on Escape
    document.addEventListener('keydown', onEscKey);

    // Listen for messages from iframe
    window.addEventListener('message', onIframeMessage);
  }

  function closeModal() {
    if (modalOverlay && modalOverlay.parentNode) {
      modalOverlay.parentNode.removeChild(modalOverlay);
    }
    modalOverlay = null;
    document.removeEventListener('keydown', onEscKey);
    window.removeEventListener('message', onIframeMessage);
  }

  function onEscKey(e) {
    if (e.key === 'Escape') closeModal();
  }

  function onIframeMessage(e) {
    if (e.data === 'sop-close-modal') {
      closeModal();
      return;
    }

    // Handle generate request from iframe
    if (e.data && e.data.type === 'sop-generate') {
      handleGenerateDownload(e.data.session_id, e.data.body, e.source);
    }
  }

  function handleGenerateDownload(sessionId, formBody, iframeWindow) {
    fetch(SOP_BACKEND + '/api/generate-and-download/' + sessionId, {
      method: 'POST',
      headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
      body: formBody
    })
    .then(function(response) {
      if (!response.ok) throw new Error('Server error');
      var disposition = response.headers.get('Content-Disposition') || '';
      var match = disposition.match(/filename="?([^"]+)"?/);
      var filename = match ? match[1] : 'SOP_Document.docx';
      return response.blob().then(function(blob) {
        return { blob: blob, filename: filename };
      });
    })
    .then(function(result) {
      var url = URL.createObjectURL(result.blob);
      var a = document.createElement('a');
      a.href = url;
      a.download = result.filename;
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);

      if (iframeWindow) {
        try { iframeWindow.postMessage('sop-download-complete', '*'); } catch(e) {}
      }
      // Close modal after short delay
      setTimeout(closeModal, 500);
    })
    .catch(function(err) {
      console.error('[SOP Generator] Download error:', err);
      if (iframeWindow) {
        try { iframeWindow.postMessage('sop-download-error', '*'); } catch(e) {}
      }
      window.alert('SOP Generator Error: ' + err.message);
    });
  }

  // ============================================================
  // Generate SOP function - called by React component
  // ============================================================

  function generateSOP() {
    if (!bpmnjsInstance) {
      window.alert('SOP Generator Error: BPMN modeler not available.');
      return;
    }

    bpmnjsInstance.saveXML({ format: true }).then(function(result) {
      var xml = result.xml;

      postWithRetry(SOP_BACKEND + '/api/upload-xml', { xml: xml }, 0)
        .then(function(response) {
          if (!response.ok) {
            return response.json().then(function(err) {
              throw new Error(err.error || 'Server error');
            });
          }
          return response.json();
        })
        .then(function(data) {
          openModal(SOP_BACKEND + '/preview/' + data.session_id);
        })
        .catch(function(err) {
          console.error('[SOP Generator] Error:', err);
          if (err.message && err.message.indexOf('fetch') !== -1) {
            window.alert(
              'SOP Generator: Backend server not available.\n\n' +
              'The SOP Generator backend may not be running. Please restart Camunda Modeler.'
            );
          } else {
            window.alert('SOP Generator Error: ' + err.message);
          }
        });
    }).catch(function(err) {
      console.error('[SOP Generator] Failed to export XML:', err);
      window.alert('SOP Generator Error: Could not export diagram XML.\n' + err.message);
    });
  }

  // Store function globally for React component
  generateSopFunction = generateSOP;

  // ============================================================
  // bpmn-js Module - registers to get bpmnjs instance
  // ============================================================

  function SopEditorAction(bpmnjs, editorActions) {
    // Store bpmnjs instance for React component to use
    bpmnjsInstance = bpmnjs;

    // Also register menu action for keyboard shortcut compatibility
    editorActions.register('generate-sop', function() {
      generateSOP();
    });
  }

  SopEditorAction.$inject = ['bpmnjs', 'editorActions'];

  var sopModule = {
    __init__: ['sopEditorAction'],
    sopEditorAction: ['type', SopEditorAction]
  };

  // ============================================================
  // React Component - EXPORT button
  // ============================================================

  // Add EXPORT button to status bar
  function addExportButton() {
    // Wait for React and Fill to be available
    if (!window.react || !window.components || !window.components.Fill) {
      setTimeout(addExportButton, 100);
      return;
    }

    const React = window.react;
    const { Fill } = window.components;

    function SopGeneratorButton(props) {
      const [isBpmnTab, setIsBpmnTab] = React.useState(false);

      React.useEffect(function() {
        // Check if BPMN tab is active by looking for bpmn-js elements
        function checkTab() {
          const canvas = document.querySelector('.djs-container, [data-container-id="canvas"]');
          const isBpmn = canvas !== null;
          setIsBpmnTab(isBpmn);
        }

        checkTab();
        const interval = setInterval(checkTab, 500);
        window.addEventListener('focus', checkTab);

        return function() {
          clearInterval(interval);
          window.removeEventListener('focus', checkTab);
        };
      }, []);

      function handleClick() {
        if (generateSopFunction) {
          generateSopFunction();
        } else {
          window.alert('SOP Generator: BPMN modeler not ready. Please wait a moment and try again.');
        }
      }

      if (!isBpmnTab) {
        return null;
      }

      return React.createElement(Fill, {
        slot: 'status-bar__file',
        group: '9_sop-export'
      }, React.createElement('button', {
        onClick: handleClick,
        title: 'Export SOP Document',
        className: 'btn',
        style: { marginLeft: '8px' }
      }, 'EXPORT'));
    }

    // Try to register as plugin component
    try {
      var plugins = window.plugins || [];
      window.plugins = plugins;
      plugins.push({
        plugin: SopGeneratorButton,
        type: 'bpmn.modeler.plugin'
      });
    } catch (e) {
      console.log('[SOP Generator] React component registration failed, button will be available via keyboard shortcut');
    }
  }

  // Initialize button when ready
  addExportButton();

  // ============================================================
  // Register plugin
  // ============================================================

  var plugins = window.plugins || [];
  window.plugins = plugins;

  // Register bpmn-js module
  plugins.push({
    plugin: sopModule,
    type: 'bpmn.modeler.additionalModules'
  });

  // Register React component plugin
  plugins.push({
    plugin: SopGeneratorButton,
    type: 'bpmn.modeler.plugin'
  });

})();

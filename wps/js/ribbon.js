/* global wps */

function resolveTaskPaneUrl() {
  var href = window.location.href;
  var base = href.substring(0, href.lastIndexOf('/'));
  return base + '/taskpane.html';
}

function OnDeckenAddinLoad(ribbonUI) {
  if (typeof wps !== 'undefined' && wps.PluginStorage) {
    wps.PluginStorage.setItem('decken_ribbon_ready', '1');
  }
  return ribbonUI;
}

function OnToggleDeckenTaskPane() {
  if (typeof wps === 'undefined' || !wps.CreateTaskPane) {
    return;
  }
  var paneId = wps.PluginStorage.getItem('decken_taskpane_id');
  if (!paneId) {
    var taskPane = wps.CreateTaskPane(resolveTaskPaneUrl());
    taskPane.DockPosition = wps.Enum && wps.Enum.msoCTPDockPositionRight ? wps.Enum.msoCTPDockPositionRight : 2;
    taskPane.Width = 380;
    taskPane.Visible = true;
    wps.PluginStorage.setItem('decken_taskpane_id', String(taskPane.ID));
  } else {
    var existing = wps.GetTaskPane(Number(paneId));
    if (existing) {
      existing.Visible = !existing.Visible;
    }
  }
}

window.OnDeckenAddinLoad = OnDeckenAddinLoad;
window.OnToggleDeckenTaskPane = OnToggleDeckenTaskPane;

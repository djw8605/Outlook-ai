/* global Office */

Office.onReady(() => {
  // Commands module is loaded — nothing to initialize here.
  // Individual ribbon button callbacks are registered below.
});

/**
 * Ribbon button handler: opens the task pane.
 * The manifest wires this function to the ShowTaskpane action.
 */
function openTaskPane(event: Office.AddinCommands.Event): void {
  event.completed();
}

// Make the function available in the global scope so the manifest can reference it.
(globalThis as Record<string, unknown>)["openTaskPane"] = openTaskPane;

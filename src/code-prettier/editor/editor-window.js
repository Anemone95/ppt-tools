// Code Prettier editor window for PPT Tools.
// Usage: osascript -l JavaScript editor-window.js <editor-dir> <state>
// <state> and the returned result use one "key=value" record per line;
// values escape backslash, newline, carriage return and tab as \\ \n \r \t.

ObjC.import("Cocoa");
ObjC.import("WebKit");

const WINDOW_STYLE = 1 | 2 | 4 | 8; // titled, closable, miniaturizable, resizable
const BACKING_STORE_BUFFERED = 2;
const VIEW_SIZABLE = 2 | 16; // width and height sizable
const INJECT_AT_DOCUMENT_START = 0;
const ACTIVATION_POLICY_REGULAR = 0;
const ANY_EVENT_MASK = -1; // NSUIntegerMax once bridged

let finished = false;
let resultText = "status=cancel";

function escapeValue(value) {
  return String(value)
    .replaceAll("\\", "\\\\")
    .replaceAll("\n", "\\n")
    .replaceAll("\r", "\\r")
    .replaceAll("\t", "\\t");
}

function unescapeValue(value) {
  const replacements = { "\\": "\\", n: "\n", r: "\r", t: "\t" };
  return value.replace(/\\(.)/g, (match, code) => replacements[code] ?? code);
}

function parseState(text) {
  const state = {};
  text
    .replace(/\r\n?/g, "\n")
    .split("\n")
    .forEach((line) => {
      const separator = line.indexOf("=");
      if (separator > 0) {
        state[line.slice(0, separator)] = unescapeValue(line.slice(separator + 1));
      }
    });
  return state;
}

function serializeResult(result) {
  const records = [
    "status=ok",
    `language=${escapeValue(result.language)}`,
    `fontSize=${result.fontSize}`,
    `lineNumbers=${result.lineNumbers ? 1 : 0}`,
    `numberGap=${result.numberGap}`,
    `defaultColor=${result.defaultColor}`,
    `lineHighlights=${result.lineHighlights}`,
    `source=${escapeValue(result.source)}`,
    `text=${escapeValue(result.text)}`,
  ];
  result.runs.forEach((run) => {
    records.push(`run=${run.start},${run.length},${run.color},${run.bold ? 1 : 0},${run.italic ? 1 : 0}`);
  });
  result.highlights.forEach((highlight) => {
    records.push(`highlight=${highlight.start},${highlight.length},${highlight.color}`);
  });
  return records.join("\n");
}

function finish(text) {
  if (finished) return;
  finished = true;
  if (text) resultText = text;
}

function buildMainMenu(app) {
  const mainMenu = $.NSMenu.alloc.init;

  const appItem = $.NSMenuItem.alloc.init;
  const appMenu = $.NSMenu.alloc.initWithTitle("Code Prettier");
  appMenu.addItemWithTitleActionKeyEquivalent("Close Code Prettier", "performClose:", "w");
  appItem.submenu = appMenu;
  mainMenu.addItem(appItem);

  const editItem = $.NSMenuItem.alloc.init;
  const editMenu = $.NSMenu.alloc.initWithTitle("Edit");
  editMenu.addItemWithTitleActionKeyEquivalent("Undo", "undo:", "z");
  editMenu.addItemWithTitleActionKeyEquivalent("Redo", "redo:", "Z");
  editMenu.addItem($.NSMenuItem.separatorItem);
  editMenu.addItemWithTitleActionKeyEquivalent("Cut", "cut:", "x");
  editMenu.addItemWithTitleActionKeyEquivalent("Copy", "copy:", "c");
  editMenu.addItemWithTitleActionKeyEquivalent("Paste", "paste:", "v");
  editMenu.addItemWithTitleActionKeyEquivalent("Select All", "selectAll:", "a");
  editItem.submenu = editMenu;
  mainMenu.addItem(editItem);

  app.mainMenu = mainMenu;
}

ObjC.registerSubclass({
  name: "PPTToolsCodePrettierController",
  superclass: "NSObject",
  methods: {
    "userContentController:didReceiveScriptMessage:": {
      types: ["void", ["id", "id"]],
      implementation: function (controller, message) {
        const payload = JSON.parse(ObjC.unwrap(message.body));
        finish(payload.action === "apply" ? serializeResult(payload) : null);
      },
    },
    "windowWillClose:": {
      types: ["void", ["id"]],
      implementation: function (notification) {
        finish(null);
      },
    },
  },
});

function run(argv) {
  const editorDir = argv[0];
  const state = parseState(argv[1] || "");
  const app = $.NSApplication.sharedApplication;
  const controller = $.PPTToolsCodePrettierController.alloc.init;

  app.setActivationPolicy(ACTIVATION_POLICY_REGULAR);
  buildMainMenu(app);
  app.finishLaunching;

  const configuration = $.WKWebViewConfiguration.alloc.init;
  const contentController = configuration.userContentController;
  contentController.addScriptMessageHandlerName(controller, "pptTools");
  contentController.addUserScript(
    $.WKUserScript.alloc.initWithSourceInjectionTimeForMainFrameOnly(
      `window.PPT_TOOLS_STATE = ${JSON.stringify(state)};`,
      INJECT_AT_DOCUMENT_START,
      true,
    ),
  );

  const window = $.NSWindow.alloc.initWithContentRectStyleMaskBackingDefer(
    $.NSMakeRect(0, 0, 1240, 760),
    WINDOW_STYLE,
    BACKING_STORE_BUFFERED,
    false,
  );
  window.title = state.mode === "edit" ? "Code Prettier · Edit code" : "Code Prettier";
  window.minSize = $.NSMakeSize(900, 560);
  window.releasedWhenClosed = false;
  window.delegate = controller;

  const webView = $.WKWebView.alloc.initWithFrameConfiguration(
    window.contentView.bounds,
    configuration,
  );
  webView.autoresizingMask = VIEW_SIZABLE;
  window.contentView.addSubview(webView);

  const editorUrl = $.NSURL.fileURLWithPath(`${editorDir}/editor.html`);
  webView.loadFileURLAllowingReadAccessToURL(editorUrl, $.NSURL.fileURLWithPath(editorDir));

  window.center;
  window.makeKeyAndOrderFront(null);
  window.makeFirstResponder(webView);
  app.activateIgnoringOtherApps(true);

  while (!finished) {
    const event = app.nextEventMatchingMaskUntilDateInModeDequeue(
      ANY_EVENT_MASK,
      $.NSDate.dateWithTimeIntervalSinceNow(0.1),
      $.NSDefaultRunLoopMode,
      true,
    );
    if (!event.isNil()) app.sendEvent(event);
  }

  window.delegate = null;
  window.orderOut(null);
  return resultText;
}

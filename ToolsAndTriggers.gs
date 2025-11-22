function onOpen() {
  const ui = SpreadsheetApp.getUi();

  ui.createMenu("Automation Tools")
    .addItem("Generate Letter", "showLetterGeneratorModal")
    // .addSeparator()
    // .addItem("Debug Tools", "showDebugMenu")
    .addToUi();

  setTimeout(() => warmCaches(), 1000);
}
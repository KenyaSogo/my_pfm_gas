SpreadsheetApp.getActiveSpreadsheet = jest.fn().mockImplementation(
  () => {
    return new MockSpreadSheet();
  }
)

class MockSpreadSheet {
  getSheetByName(name){
    switch(name){
      case "settings":
        return new MockSettingsSheet(name);
      default:
        return "default sheet hoge";
    }
  }
}

class MockSettingsSheet {
  constructor(name){
    this.name = name;
  }
  getRange(name) {
    switch(name){
      case "phantom_js_key":
        return new MockPhantomJsKeyRange();
    }
  }
}

class MockPhantomJsKeyRange {
  getValue() {
    return "mock-phantom-js-key";
  }
}

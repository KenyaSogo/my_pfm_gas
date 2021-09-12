require("./setup.js");
const otameshi = require("../otameshi.js");

describe('hoge', () => {
  it("fuga", () => {
    expect(
      otameshi()
    ).toEqual("mock-phantom-js-key");
  });
})

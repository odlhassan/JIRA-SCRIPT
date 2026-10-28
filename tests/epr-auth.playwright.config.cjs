module.exports = {
  testDir: __dirname, testMatch: 'test_epr_auth_browser.cjs', timeout: 30000,
  use: { browserName: 'chromium', viewport: {width:1440,height:1000} },
};

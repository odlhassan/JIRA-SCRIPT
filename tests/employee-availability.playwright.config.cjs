module.exports = {
  testDir: __dirname,
  testMatch: 'test_employee_availability_calendar.cjs',
  timeout: 30000,
  use: { browserName: 'chromium', viewport: { width: 1440, height: 1000 } },
};

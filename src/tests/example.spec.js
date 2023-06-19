// @ts-check
const { expect } = require('@playwright/test');
const { test } = require('./fixtures/test-fixtures');
const { dynamicsUrl } = require('./test-config');

const expectedWrongUrlMessage = 'Dynamics environment needs to be open in the current tab';
const expectedSuccessMessage = '204 : Success, Record(s) were added for you';

test('Dynamics URL not found warning', async ({ page, extensionId }) => {
  await page.goto(`chrome-extension://${extensionId}/popup.html`);
  expect(await page.locator('.warning', { hasText: expectedWrongUrlMessage }).isVisible());
});

test('Can add a Contact record', async ({ page, extensionId }) => {
  await page.goto(`chrome-extension://${extensionId}/popup.html`);
  await page.getByLabel('Environment:').type(dynamicsUrl);
  await page.getByRole('button', { name: 'Add' }).click();

  const loadingSpinner = page.getByAltText('loading-spinner');
  await loadingSpinner.waitFor({state: 'detached'});

  const successMessage = page.getByText(expectedSuccessMessage);
  await expect(successMessage).toBeVisible();
});

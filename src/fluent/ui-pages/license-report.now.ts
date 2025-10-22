import '@servicenow/sdk/global';
import { UiPage } from '@servicenow/sdk/core';
import reportPage from '../../client/index.html';

export const license_report_dashboard = UiPage({
  $id: Now.ID['license-report-dashboard'],
  endpoint: 'x_hete_license_6h4_report.do',
  html: reportPage,
  direct: true
});

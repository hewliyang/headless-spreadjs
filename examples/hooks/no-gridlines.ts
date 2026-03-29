/**
 * Hide gridlines on every sheet when a workbook is opened.
 *
 * Install:
 *   mkdir -p .headless-spreadjs/hooks
 *   cp examples/hooks/no-gridlines.ts .headless-spreadjs/hooks/
 */

import type { HookAPI, HookContext } from "@hewliyang/headless-spreadjs/hooks";

export default function (hsx: HookAPI) {
  hsx.on("onOpen", function hideGridlines(ctx: HookContext) {
    const count = ctx.workbook.getSheetCount();
    let changed = 0;

    for (let i = 0; i < count; i++) {
      const sheet = ctx.workbook.getSheet(i);
      const gl = sheet.options.gridline;

      if (
        gl?.showVerticalGridline === false &&
        gl?.showHorizontalGridline === false
      ) {
        continue;
      }

      sheet.options.gridline = {
        showVerticalGridline: false,
        showHorizontalGridline: false,
      };
      changed++;
    }

    if (changed > 0) {
      console.log(`Hidden gridlines on ${changed}/${count} sheet(s)`);
    }
  });
}

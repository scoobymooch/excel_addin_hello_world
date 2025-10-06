Office.onReady(() => {
  Excel.run(async context => {
    context.workbook.worksheets.onSelectionChanged.add(onSelectionChanged);
    await context.sync();
  });
});

async function onSelectionChanged(eventArgs) {
  await Excel.run(async context => {
    const range = context.workbook.getSelectedRange();
    range.load("address, values");
    await context.sync();
    document.getElementById("cellData").innerText =
      `Address: ${range.address}\nValue: ${range.values[0][0]}`;
  });
}
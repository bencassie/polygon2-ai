/* global console, document, Excel, Office */

Office.onReady(() => {
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("run").onclick = run;
  document.getElementById("setApiKey").onclick = setApiKey;

  // Prefill API key if saved
  const savedKey = localStorage.getItem("polygonApiKey");
  if (savedKey) {
    document.getElementById("apiKeyInput").value = savedKey;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      const activeCell = context.workbook.getActiveCell();
      activeCell.formulas = [["=POLYGON.getTickerDetails(\"AAPL\")"]];
      await context.sync();
      console.log("Inserted example formula in active cell.");
    });
  } catch (error) {
    console.error(error);
  }
}

let apiKey = "";

export async function setApiKey() {
  try {
    await Excel.run(async (context) => {
      const input = document.getElementById("apiKeyInput").value;
      apiKey = input;
      localStorage.setItem("polygonApiKey", input);
      document.getElementById("apiKeyStatus").textContent = "API key saved successfully.";
      console.log("API key saved!");
    });
  } catch (error) {
    console.error(error);
    document.getElementById("apiKeyStatus").textContent = "Error saving API key.";
  }
}
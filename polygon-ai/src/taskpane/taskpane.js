Office.onReady(() => {
  // Hide sideload message and show app body
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";
  document.getElementById("setApiKey").onclick = setApiKey;

  // Prefill API key if saved
  const savedKey = localStorage.getItem("polygonApiKey");
  if (savedKey) {
    document.getElementById("apiKeyInput").value = savedKey;
  }

  // Define custom functions with name, description, and example formula
  const customFunctions = [
    {
      name: "getTickerDetails",
      description: "Gets a specific ticker detail from Polygon.io.",
      example: '=POLYGON.getTickerDetails("AAPL")'
    },
    {
      name: "getTickerNews",
      description: "Retrieves news articles for a specific stock ticker from Polygon.io.",
      example: '=POLYGON.getTickerNews("AAPL")'
    }
    // Add more custom functions here as needed
  ];

  // Populate the functions list
  const functionsList = document.getElementById("functions-list");
  customFunctions.forEach(func => {
    const functionDiv = document.createElement("div");
    functionDiv.className = "function";

    const nameHeader = document.createElement("h3");
    nameHeader.textContent = func.name;
    functionDiv.appendChild(nameHeader);

    const descPara = document.createElement("p");
    descPara.textContent = func.description;
    functionDiv.appendChild(descPara);

    const insertButton = document.createElement("button");
    insertButton.textContent = "Insert";
    insertButton.onclick = () => insertFunction(func.example);
    functionDiv.appendChild(insertButton);

    functionsList.appendChild(functionDiv);
  });
});

// Function to insert a formula into the active cell
async function insertFunction(formula) {
  try {
    await Excel.run(async (context) => {
      const activeCell = context.workbook.getActiveCell();
      activeCell.formulas = [[formula]];
      await context.sync();
    });
  } catch (error) {
    console.error("Error inserting formula:", error);
  }
}

// Function to save the API key
async function setApiKey() {
  try {
    await Excel.run(async () => {
      const input = document.getElementById("apiKeyInput").value;
      localStorage.setItem("polygonApiKey", input);
      document.getElementById("apiKeyStatus").textContent = "API key saved successfully.";
      console.log("API key saved!");
    });
  } catch (error) {
    console.error("Error saving API key:", error);
    document.getElementById("apiKeyStatus").textContent = "Error saving API key.";
  }
}
Office.onReady(() => {
  // Hide sideload message and show app body
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  // Define custom functions with signature and example formula
  const customFunctions = [
      {
          signature: "=POLYGON.getTickerDetails(ticker, [property])",
          example: '=POLYGON.getTickerDetails("AAPL")'
      },
      {
          signature: "=POLYGON.getTickerNews(ticker, [limit])",
          example: '=POLYGON.getTickerNews("AAPL")'
      }
  ];

  // Populate the functions list
  const functionsList = document.getElementById("functions-list");
  customFunctions.forEach(func => {
      const functionDiv = document.createElement("div");
      functionDiv.className = "function";

      const signatureHeader = document.createElement("h3");
      signatureHeader.textContent = func.signature;
      functionDiv.appendChild(signatureHeader);

      const insertButton = document.createElement("button");
      insertButton.textContent = "Insert";
      insertButton.onclick = () => insertFunction(func.example);
      functionDiv.appendChild(insertButton);

      functionsList.appendChild(functionDiv);
  });

  // API key handling
  const savedKey = localStorage.getItem("polygonApiKey");
  toggleApiKeyInput(!savedKey);

  document.getElementById("setApiKey").onclick = setApiKey;
  document.getElementById("changeApiKey").onclick = changeApiKey;
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

// Function to toggle API key input visibility
function toggleApiKeyInput(showInput) {
  const inputContainer = document.getElementById("apiKeyInputContainer");
  const apiKeySet = document.getElementById("apiKeySet");
  if (showInput) {
      inputContainer.style.display = "block";
      apiKeySet.style.display = "none";
  } else {
      inputContainer.style.display = "none";
      apiKeySet.style.display = "block";
  }
}

// Function to save the API key
function setApiKey() {
  const input = document.getElementById("apiKeyInput").value;
  localStorage.setItem("polygonApiKey", input);
  document.getElementById("apiKeyStatus").textContent = "API key saved successfully.";
  toggleApiKeyInput(!input);
}

// Function to change the API key
function changeApiKey() {
  document.getElementById("apiKeyInput").value = "";
  toggleApiKeyInput(true);
}
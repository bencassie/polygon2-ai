Office.onReady(() => {
  // Hide sideload message and show app body
  document.getElementById("sideload-msg").style.display = "none";
  document.getElementById("app-body").style.display = "flex";

  // Define custom functions with signature and example formula
  const customFunctions = [
    {
      signature: "getDividends(ticker, [limit])",
      example: '=POLYGON.getDividends("AAPL", 4)',
      description: "Retrieves dividend information for a ticker"
    },
    {
      signature: "getExchanges([exchangeType])",
      example: '=POLYGON.getExchanges("exchange")',
      description: "Retrieves information about stock exchanges"
    },
    {
      signature: "getHistoricalOHLC(ticker, fromDate, toDate, [timespan])",
      example: '=POLYGON.getHistoricalOHLC("AAPL", "2023-01-01", "2023-12-31", "day")',
      description: "Retrieves historical OHLC data for a specific date range"
    },
    {
      signature: "getLatestPrice(ticker, [property])",
      example: '=POLYGON.getLatestPrice("AAPL")',
      description: "Retrieves the latest stock price data for a ticker"
    },
    {
      signature: "getMarketHolidays([year])",
      example: '=POLYGON.getMarketHolidays(2025)',
      description: "Gets upcoming market holidays and trading hours"
    },
    {
      signature: "getMarketStatus([market])",
      example: '=POLYGON.getMarketStatus("us")',
      description: "Gets the current market status from Polygon.io"
    },
    {
      signature: "getPortfolioSummary(portfolioData)",
      example: '=POLYGON.getPortfolioSummary(A1:B10)',
      description: "Creates a simple portfolio tracker with current values and returns"
    },
    {
      signature: "getSectorPerformance([timespan])",
      example: '=POLYGON.getSectorPerformance("week")',
      description: "Analyzes and compares performance of major market sectors"
    },
    {
      signature: "getStockCorrelation(ticker1, ticker2, [days])",
      example: '=POLYGON.getStockCorrelation("AAPL", "MSFT", 30)',
      description: "Calculates correlation between two stocks over a period"
    },
    {
      signature: "getStockSplits(ticker, [limit])",
      example: '=POLYGON.getStockSplits("AAPL", 5)',
      description: "Retrieves stock splits history for a ticker"
    },
    {
      signature: "getTechnicalIndicator(ticker, indicator, [period], [from], [to])",
      example: '=POLYGON.getTechnicalIndicator("AAPL", "SMA", 14)',
      description: "Calculates technical indicators for a stock"
    },
    {
      signature: "getTickerDetails(ticker, [property])",
      example: '=POLYGON.getTickerDetails("AAPL")',
      description: "Gets details for a specific ticker from Polygon.io"
    },
    {
      signature: "getTickerNews(ticker, [limit])",
      example: '=POLYGON.getTickerNews("AAPL", 5)',
      description: "Retrieves news articles for a specific stock ticker"
    },
    {
      signature: "searchTickers(searchTerm, [limit])",
      example: '=POLYGON.searchTickers("Apple", 10)',
      description: "Searches for tickers matching a search term"
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